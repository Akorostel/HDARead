using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NDesk.Options;

using Opc;
using OpcCom;

using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

using System.IO;

/* TODO:
 * +aggregate enum
 * +resample interval
 * +file output
 * +read raw
 * +input file with tag list
 * + why we use OPCTrend and not OPCServer read method
 * + maybe its better to query data tag by tag (because 'tag not found' exception doesn't show which tag is wrong)
 * +   or maybe validate tags before read and delete them from query?
 * +command line parameter to display output quality or not
 * +command line parameter to 'merged' output format (single timestamp column for all tags)
 * + OutputTable: What if different tags have different number of points?!
 * + header is shifted by one column in TABLE mode
 * + option to toggle debug output
 * + check if 'MaxValues' work. If not - delete it.
 * + trying to get data for the whole month - E_MAXEXCEEDED
 * remember that timestamps may be in reverse order. Currently this will crash 'Merge' algorithm and may be something else
 * better merge algorithm?
 * exceptions
 * if output file already exists: overwrite or do nothing
 * omit value if NODATA quality
 * shutdown event
 * log to file?
 * partition the program to separate modules/classes...
 * write to file portion by portion to conserve memory...
 */

namespace HDARead {

    enum eOutputFormat {
        LIST = 1,
        TABLE = 2,
        MERGED = 3
    }
    enum eOutputQuality {
        NONE = 0,
        DA = 1,
        HISTORIAN = 2,
        BOTH = 3
    }
    class Program {
        static string Host = null;
        static string Server = null;
        static string StartTime = "NOW-1H";
        static string EndTime = "NOW";
        static int Aggregate = (int)HDAClient.OPCHDA_AGGREGATE.AVERAGE;
        static string OutputTimestampFormat = null;
        static int MaxValues = int.MaxValue; 
        static int ResampleInterval = 0;
        static bool IncludeBounds = false; // 
        static string OutputFileName = null;
        static string InputFileName = null;
        static bool ReadRaw = false;
        static bool Help = false;
        static bool Verbose = false;
        static eOutputQuality OutputQuality = eOutputQuality.NONE;
        static eOutputFormat OutputFormat = eOutputFormat.LIST;
        static List<string> Tagnames = new List<string>();
        static string OptionDescription = null;

        static TraceSource _trace = new TraceSource("ConsoleApplicationTraceSource");

        static void Main(string[] args) {

            if (!ParseCommandLine(args)) {
                return;
            }
            if (!CheckOptions()) {
                return;
            }
            ShowInfo();

            var srv = new HDAClient(_trace.Switch);

            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;
            try {
                bool res = false;
                if (srv.Connect(Host, Server)) {
                    // Remove unknown items from the list
                    srv.Validate(Tagnames);
                    // Read items using the Hda.Server class.
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, IncludeBounds, ReadRaw, out OPCHDAItemValues);
                    //res = srv.ReadTrend(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, IncludeBounds, ReadRaw, out OPCHDAItemValues);
                } else {
                    Console.WriteLine("HDARead unable to connect to OPC server.");
                }
                if (!res) {
                    Console.WriteLine("HDARead Error.");
                    return;
                }
            } catch (Exception e) {
                Console.WriteLine(e.Message);
                return;
            } finally {
                srv.Disconnect();
            }

            if (OPCHDAItemValues == null) {
                Console.WriteLine("HDARead returned null.");
                return;
            } else {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("HDARead OK.");
                Console.ResetColor();
            }

            WriteOutput(OPCHDAItemValues);
           
            return;
        }

        // If parsing was unsuccessfull, return false
        static bool ParseCommandLine(string[] args) {
             var p = new OptionSet() {
   	            { "n=|node=",               "Remote computer name (optional)",  v => Host = v },
   	            { "s=|server=",             "OPC HDA server name (required)",   v => Server = v },
   	            { "from=|start=|begin=",    "Start time (abs. or relative), default NOW-1H",    
                                                                                v => StartTime = v ?? "NOW-1H"},
   	            { "to=|end=",               "End time (abs. or relative), default NOW",      
                                                                                v => EndTime = v ?? "NOW"},
   	            { "a=|agg=",                "Aggregate (see spec)",             v => Aggregate = GetHDAAggregate(v) },
                { "r=|resample=",           "Resample interval (in seconds), 0 - return just one value (see OPC HDA spec.)",  
                                                                                v => ResampleInterval = Int32.Parse(v)},
                { "raw",                    "Read raw data (if omitted, read processed data) ",  
                                                                                v => ReadRaw = v != null},
                { "m=|maxvalues=",          "Maximum number of values to load (only for raw data)", 
                                                                                v => MaxValues = Int32.Parse(v)},
                { "b|bounds",               "Whether the bounding item values should be returned (only for ReadRaw).",  
                                                                                v => IncludeBounds = v != null},
                { "t=|tsformat=",           "Output timestamp format to use",   v => OutputTimestampFormat = v},
                { "f=",                     "Output format (LIST or TABLE or MERGED)",   
                                                                                v => OutputFormat = GetOutputFormat(v)},
                { "q=",                     "Include quality in output data (NONE, DA, HISTORIAN or BOTH)",   
                                                                                v => OutputQuality = GetOutputQuality(v)},
                { "o=|output=",             "Output filename (if omitted, output to console)",   
                                                                                v => OutputFileName = v},
                { "i=|input=",              "Input filename with list of tags (if omitted, tag list must be provided as command line argument)",   
                                                                                v => InputFileName = v},
                { "v|verbose",              "Output debug info",                v => Verbose = v != null},
                { "h|?|help",               "Show help",                        v => Help = v != null},
                { "<>",                     "List of tag names",                v => Tagnames.Add (v)},
            };

            Console.Write("HDARead: ");
            try {
                p.Parse(args);
            } catch (OptionException e) {
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `HDARead --help' for more information.");
                return false;
            }

            StringBuilder sb = new StringBuilder();
            TextWriter tw = new StringWriter(sb);
            if(tw != null) {
                p.WriteOptionDescriptions(tw);
                tw.Flush();
                OptionDescription = sb.ToString();
            }
            return true;
        }


        static bool CheckOptions() {
            if (Help) {
                ShowHelp();
                return false;
            }

            if (Verbose) {
                _trace.Switch.Level = SourceLevels.All;
            } else {
                _trace.Switch.Level = SourceLevels.Information;
            }

            if (string.IsNullOrEmpty(Server)) {
                Console.WriteLine("Missing required option s=|server=");
                return false;
            }

            if (string.IsNullOrEmpty(InputFileName)) {
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return false;
                }
            } else {
                if (Tagnames.Count() > 0) {
                    Console.WriteLine("If the input file is specified, no tags may be entered as command line argument");
                    return false;
                }
                // try catch !!!
                Tagnames = File.ReadLines(InputFileName).ToList();
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return false;
                }
            }

            return true;
        }

        static void ShowInfo() {
            Console.WriteLine("HDARead is going to:");
            Console.WriteLine("\t connect to OPC HDA server named '{0}' on computer '{1}'", Server, Host);
            Console.WriteLine("\t and read {0} data for the period from '{1}' to '{2}'", ReadRaw ? "raw" : "processed", StartTime, EndTime);
            if (!ReadRaw) {
                Console.WriteLine("\t aggregating as {0}", ((HDAClient.OPCHDA_AGGREGATE)Aggregate).ToString());
                Console.WriteLine("\t with resample interval {0} seconds.", ResampleInterval);
            } else {
                Console.WriteLine("\t No more than {0} values for each tag should be loaded.", MaxValues);
                Console.WriteLine("\t Bounding item values will {0} be included in result.", IncludeBounds ? "" : "not");
            }
            if (!string.IsNullOrEmpty(OutputTimestampFormat)) 
                Console.WriteLine("\t Output timestamp format is specified as {0}.", OutputTimestampFormat);
            if (!string.IsNullOrEmpty(OutputFileName))
                Console.WriteLine("\t The resulting data will be written to file {0}.", OutputFileName);
            else
                Console.WriteLine("\t The resulting data will be output to console.", OutputFileName);
            Console.WriteLine("\t The list of requested tags:");
            foreach (string t in Tagnames) {
                Console.WriteLine("\t\t" + t);
            }

        }

        static bool WriteOutput(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            StreamWriter writer=null;
            try {
                if (!string.IsNullOrEmpty(OutputFileName)) {
                    writer = new StreamWriter(OutputFileName);
                } else {
                    writer = new StreamWriter(Console.OpenStandardOutput());
                    writer.AutoFlush = true;
                    Console.SetOut(writer);
                }

                switch (OutputFormat) {
                    case eOutputFormat.LIST:
                        OutputList(writer, OPCHDAItemValues);
                        break;
                    case eOutputFormat.MERGED:
                        if (ReadRaw)
                            OutputMerged(writer, Merge(OPCHDAItemValues));
                        else
                            OutputMerged(writer, OPCHDAItemValues);
                        break;
                    case eOutputFormat.TABLE:
                        OutputTable(writer, OPCHDAItemValues);
                        break;
                }

                if (!string.IsNullOrEmpty(OutputFileName)) {
                    Console.WriteLine("Data were written to file {0}.", OutputFileName);
                }
                return true;
            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Opc.ResultIDException:" + e.ToString());
                return false;
            } finally {
                if (!string.IsNullOrEmpty(OutputFileName) && (writer!=null) ) {
                    writer.Close();
                }
            }
        }

        // Output format: list
        // tag1 timestamp, tag1 value
        // ...
        // tag2 timestamp, tag2 value
        // ...
        static void OutputList(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                sw.WriteLine();
                sw.WriteLine("\tTag ({0} of {1}): {2} ({3} values):", i + 1, 
                    OPCHDAItemValues.Count(), 
                    OPCHDAItemValues[i].ItemName, 
                    OPCHDAItemValues[i].Count);

                // header
                string valstr = "{0,20},{1,20}";
                sw.WriteLine("{0,20}{1,20}", "Timestamp", "Value");
                if ((OutputQuality == eOutputQuality.DA) || (OutputQuality == eOutputQuality.BOTH)) {
                    sw.WriteLine("{0,20}", "DA Quality");
                    valstr += "{2,20}";
                }
                if ((OutputQuality == eOutputQuality.HISTORIAN) || (OutputQuality == eOutputQuality.BOTH)) {
                    sw.WriteLine("{0,20}", "Historian Quality");
                    valstr += "{3,20}";
                }
                // body
                for (int j = 0; j < OPCHDAItemValues[i].Count; j++) {
                    sw.WriteLine(valstr, 
                        GetDatetimeStr(OPCHDAItemValues[i][j].Timestamp, OutputTimestampFormat), 
                        OPCHDAItemValues[i][j].Value.ToString(), 
                        OPCHDAItemValues[i][j].Quality.ToString(),
                        OPCHDAItemValues[i][j].HistorianQuality.ToString());
                }
            }
        }

        // Output format: table
        // tag1 timestamp, tag1 value, tag2 timestamp, tag2 value, ...
        static void OutputTable(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {

            string hdr = "Timestamp, {0}";
            string valstr = "{0},{1}";
            string emptystr = ",";

            if ((OutputQuality == eOutputQuality.DA) || (OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} da quality";
                valstr += ",{2}";
                emptystr += ",";
            }
            if ((OutputQuality == eOutputQuality.HISTORIAN) || (OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} hist quality";
                valstr += ",{3}";
                emptystr += ",";
            }
            // header
            sw.Write(hdr, OPCHDAItemValues[0].ItemName);
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                sw.Write(", ");
                sw.Write(hdr, OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();

            // What if different tags have different number of points?!
            int max_rows = OPCHDAItemValues.Max(x => x.Count);
            /*
            int max_rows = OPCHDAItemValues[0].Count;
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                if (max_rows < OPCHDAItemValues[i].Count) {
                    max_rows = OPCHDAItemValues[i].Count;
                }
            }
            */

            // body
            for (int j = 0; j < max_rows; j++) {
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    if (i > 0) {
                        sw.Write(", ");
                    }
                    if (j < OPCHDAItemValues[i].Count) {
                        sw.Write(valstr, 
                            GetDatetimeStr(OPCHDAItemValues[i][j].Timestamp, OutputTimestampFormat), 
                            OPCHDAItemValues[i][j].Value.ToString(), 
                            OPCHDAItemValues[i][j].Quality.ToString(), 
                            OPCHDAItemValues[i][j].HistorianQuality.ToString());
                    } else {
                        sw.Write(emptystr);
                    }
                }
                sw.WriteLine();
            }
        }

        // Output format: merged
        // timestamp, tag1 value, tag2 value, ...
        static void OutputMerged(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            // header
            sw.Write("Timestamp");
            string hdr = ", {0}";
            if ((OutputQuality == eOutputQuality.DA) || (OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} da quality";
            }
            if ((OutputQuality == eOutputQuality.HISTORIAN) || (OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} hist quality";
            }
            // header
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                sw.Write(hdr, OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();
            for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                sw.Write("{0}", GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, OutputTimestampFormat));
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    // Maybe its better to catch exception (null ref) than to check every element
                    if (OPCHDAItemValues[i][j].Value == null) {
                        sw.Write(", ");
                    } else {
                        sw.Write(", {0}", OPCHDAItemValues[i][j].Value.ToString());
                    }

                    if ((OutputQuality == eOutputQuality.DA) || (OutputQuality == eOutputQuality.BOTH)) {
                        if (OPCHDAItemValues[i][j].Quality == null) {
                            sw.Write(", ");
                        } else {
                            sw.Write(", {0}", OPCHDAItemValues[i][j].Quality.ToString());
                        }
                    }

                    if ((OutputQuality == eOutputQuality.HISTORIAN) || (OutputQuality == eOutputQuality.BOTH)) {
                        // OPC.DA.Quality is struct, but OPC.HDA.Quality is enum.
                        // Enum cannot be null, so there is no need to check
                        sw.Write(", {0}", OPCHDAItemValues[i][j].HistorianQuality.ToString());
                    }
                }
                sw.WriteLine();
            }
        }

        // Merge multiple timeseries. Fill with NaN.
        static Opc.Hda.ItemValueCollection[] Merge(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            int n_tags = OPCHDAItemValues.Count();
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Starting merge. n_tags = {0}", n_tags);
            var MergedValues = new Opc.Hda.ItemValueCollection[n_tags];
            for (int i = 0; i < n_tags; i++) {
                MergedValues[i] = new Opc.Hda.ItemValueCollection(new Opc.ItemIdentifier(OPCHDAItemValues[i]));
            }

            // init pointer (row numbers) for each column
            int[] row = new int[n_tags];
            for (int i = 0; i < n_tags; i++) {
                row[i] = 0;
            }

            bool ascending = true;
            if (OPCHDAItemValues[0].EndTime < OPCHDAItemValues[0].StartTime)
                ascending = false;
       
            bool have_more_data = true;
            while (have_more_data) {

                string msg = "rows: ";
                for (int i = 0; i < n_tags; i++) {
                    msg += row[i] + ", ";
                }
                _trace.TraceEvent(TraceEventType.Verbose, 0, msg);

                // find minimum timestamp among first values of each tag (for ascending timeseries)
                // find maximum timestamp among first values of each tag (for descending timeseries)
                int ext_ts_col = -1;
                DateTime ext_ts = System.DateTime.MaxValue;

                if (ascending)
                    MinTimestamp(n_tags, row, OPCHDAItemValues, out ext_ts, out ext_ts_col);
                else
                    MaxTimestamp(n_tags, row, OPCHDAItemValues, out ext_ts, out ext_ts_col);

                have_more_data = false;
                // copy value with this timestamp to output array
                for (int i = 0; i < n_tags; i++) {
                    if ((row[i] < OPCHDAItemValues[i].Count) && (OPCHDAItemValues[i][row[i]].Timestamp.Equals(ext_ts))) {
                        MergedValues[i].Add(OPCHDAItemValues[i][row[i]]);
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Copying: {0}, {1}, {2}",
                            OPCHDAItemValues[i].ItemName,
                            OPCHDAItemValues[i][row[i]].Timestamp.ToString(),
                            OPCHDAItemValues[i][row[i]].Value.ToString());
                        row[i]++;
                        if (row[i] < OPCHDAItemValues[i].Count) {
                            have_more_data = true;
                        }
                    } else {
                        // if there is no value for this timestamp, fill blank
                        var itemvalue = new Opc.Hda.ItemValue();
                        itemvalue.Timestamp = ext_ts;
                        itemvalue.Value = null;
                        var q = new Opc.Da.Quality();
                        q.QualityBits = Opc.Da.qualityBits.uncertain;
                        itemvalue.Quality = q;
                        itemvalue.HistorianQuality = Opc.Hda.Quality.NoData;
                        MergedValues[i].Add(itemvalue);

                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Filling with blank: {0}, {1}",
                            OPCHDAItemValues[i].ItemName,
                            itemvalue.Timestamp.ToString());
                    }
                }
            }
            return MergedValues;
        }

        static void MinTimestamp(int n_tags, int[] row, Opc.Hda.ItemValueCollection[] OPCHDAItemValues,
            out DateTime ext_ts, out int ext_ts_col) {

            _trace.TraceEvent(TraceEventType.Verbose, 0, "Looking for min timestamp");
            ext_ts_col = -1;
            ext_ts = System.DateTime.MaxValue;
            for (int i = 0; i < n_tags; i++) {
                if (row[i] >= OPCHDAItemValues[i].Count) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: no data", OPCHDAItemValues[i].ItemName);
                } else {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: {1}", OPCHDAItemValues[i].ItemName, OPCHDAItemValues[i][row[i]].Timestamp);
                    if (ext_ts > OPCHDAItemValues[i][row[i]].Timestamp) {
                        ext_ts = OPCHDAItemValues[i][row[i]].Timestamp;
                        ext_ts_col = i;
                    }
                }
            }
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Min timestamp = {0}, index = {1}", ext_ts, ext_ts_col);
        
        }

        static void MaxTimestamp(int n_tags, int[] row, Opc.Hda.ItemValueCollection[] OPCHDAItemValues,
            out DateTime ext_ts, out int ext_ts_col) {

            _trace.TraceEvent(TraceEventType.Verbose, 0, "Looking for max timestamp");
            ext_ts_col = -1;
            ext_ts = System.DateTime.MinValue;
            for (int i = 0; i < n_tags; i++) {
                if (row[i] >= OPCHDAItemValues[i].Count) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: no data", OPCHDAItemValues[i].ItemName);
                } else {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: {1}", OPCHDAItemValues[i].ItemName, OPCHDAItemValues[i][row[i]].Timestamp);
                    if (ext_ts < OPCHDAItemValues[i][row[i]].Timestamp) {
                        ext_ts = OPCHDAItemValues[i][row[i]].Timestamp;
                        ext_ts_col = i;
                    }
                }
            }
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Max timestamp = {0}, index = {1}", ext_ts, ext_ts_col);

        }

        static void ShowHelp() {
            Console.WriteLine("Usage: HDARead OPTIONS tag1 tag2 tag3");
            Console.WriteLine("HDARead is used to read the data from OPC HDA server.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine(OptionDescription);
            Console.WriteLine();
            Console.WriteLine("Aggregates:");
            foreach (string agg in Enum.GetNames(typeof(HDAClient.OPCHDA_AGGREGATE))) {
                Console.Write(" " + agg);
            }
                
        }

        static int GetHDAAggregate(string str) {
            HDAClient.OPCHDA_AGGREGATE Value;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(HDAClient.OPCHDA_AGGREGATE), Value))
                return (int)Value;
            else
                throw new NDesk.Options.OptionException("Wrong aggregate: " + str, "-a");
        }

        static eOutputFormat GetOutputFormat(string str) {
            eOutputFormat Value;
            if (string.IsNullOrEmpty(str))
                return eOutputFormat.LIST;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(eOutputFormat), Value))
                return Value;
            else
                throw new NDesk.Options.OptionException("Wrong output format: " + str, "-f");
        }

        static eOutputQuality GetOutputQuality(string str) {
            eOutputQuality Value;
            if (string.IsNullOrEmpty(str))
                return eOutputQuality.NONE;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(eOutputQuality), Value))
                return Value;
            else
                throw new NDesk.Options.OptionException("Wrong output quality: " + str, "-q");
        }

        static string GetDatetimeStr(DateTime dt, string fmt) {
            string ts;
            if (string.IsNullOrEmpty(fmt)) {
                ts = dt.ToString();
            } else {
                ts = dt.ToString(fmt);
            }
            return ts;
        }

    }
}
