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
 * exceptions
 * +read raw
 * +input file with tag list
 * maybe its better to query data tag by tag (because 'tag not found' exception doesn't show which tag is wrong)
 *    or maybe validate tags before read and delete them from query?
 * command line parameter to display output quality or not
 * command line parameter to 'merged' output format (single timestamp column for all tags)
 * OutputTable: What if different tags have different number of points?!
 */

namespace HDARead {
    class Program {

        static TraceSource _trace = new TraceSource("ConsoleApplicationTraceSource");
        enum eOutputFormat {
            LIST = 1,
            TABLE = 2,
            MERGED = 3
        }

        static void Main(string[] args) {

            // Private NumberOfTags As ConInt
            string Host = null, Server = null;
            string StartTime = "NOW-1H", EndTime = "NOW";
            int Aggregate = (int)HDAClient.OPCHDA_AGGREGATE.AVERAGE;
            string OutputTimestampFormat = null;
            int MaxValues = 10;
            int ResampleInterval = 0;
            string OutputFileName = null;
            string InputFileName = null;
            bool read_raw = false;
            bool show_help = false;
            eOutputFormat OutputFormat = eOutputFormat.LIST;

            List<string> Tagnames = new List<string>();

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
                                                                                v => read_raw = v != null},
                { "m=|maxvalues=",          "Maximum number of values to load (should be checked at OPC server side, but doesn't work)", 
                                                                                v => MaxValues = Int32.Parse(v)},
                { "t=|tsformat=",           "Output timestamp format to use",   v => OutputTimestampFormat = v},
                { "f=",                     "Output format (LIST or TABLE or MERGED)",   
                                                                                v => OutputFormat = GetOutputFormat(v)},
                { "o=|output=",             "Output filename (if omitted, output to console)",   
                                                                                v => OutputFileName = v},
                { "i=|input=",              "Input filename with list of tags (if omitted, tag list must be provided as command line argument)",   
                                                                                v => InputFileName = v},
                { "h|?|help",               "Show help",                        v => show_help = v != null },
                { "<>",                     "List of tag names",                v => Tagnames.Add (v)},
            };

            Console.Write("HDARead: ");
            try {
                p.Parse(args);
            } catch (OptionException e) {
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `HDARead --help' for more information.");
                return;
            }

            if (show_help) {
                ShowHelp(p);
                return;
            }

            if (string.IsNullOrEmpty(Server)) {
                Console.WriteLine("Missing required option s=|server=");
                return;
            }

            if (string.IsNullOrEmpty(InputFileName)) {
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return;
                }
            } else {
                if (Tagnames.Count() > 0 ) {
                    Console.WriteLine("If the input file is specified, no tags may be entered as command line argument");
                    return;
                }
                Tagnames = File.ReadLines(InputFileName).ToList();
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return;
                }
            }

            Console.WriteLine("HDARead is going to:");
            Console.WriteLine("\t connect to OPC HDA server named '{0}' on computer '{1}'", Server, Host);
            Console.WriteLine("\t and read {0} data for the period from '{1}' to '{2}'", read_raw?"raw":"processed", StartTime, EndTime);
            Console.WriteLine("\t with resample interval {0} seconds", ResampleInterval, EndTime);
            Console.WriteLine("\t for the following tags:");
            foreach (string t in Tagnames) {
                Console.WriteLine("\t\t" + t );
            }
            Console.WriteLine("\t No more than {0} values should be loaded (checked only at OPC server side).", MaxValues);

            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;
            try {
                bool res = false;
                var srv = new HDAClient();
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Created HDAClient");
                if (srv.Connect(Host, Server)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Connected. Going to read.");
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, read_raw, out OPCHDAItemValues);
                } else {
                    Console.WriteLine("HDARead unable to connect to OPC server.");
                }
                srv.Disconnect();
                if (!res) {
                    Console.WriteLine("HDARead Error.");
                    return;
                }
            } catch (Exception e) {
                Console.WriteLine(e.Message);
                return;
            }
            if (OPCHDAItemValues == null) {
                Console.WriteLine("HDARead returned null.");
                return;
            } else {
                Console.WriteLine("HDARead OK.");
            }
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Number of tags = OPCHDAItemValues.Count()={0}", OPCHDAItemValues.Count());

            StreamWriter writer;
            if (!string.IsNullOrEmpty(OutputFileName)) {
                writer = new StreamWriter(OutputFileName);
            } else {
                writer = new StreamWriter(Console.OpenStandardOutput());
                writer.AutoFlush = true;
                Console.SetOut(writer);
            }

            switch (OutputFormat) {
                case eOutputFormat.LIST: 
                    OutputList(writer, OPCHDAItemValues, OutputTimestampFormat);
                    break;
                case eOutputFormat.MERGED:
                    OutputMerged(writer, OPCHDAItemValues, OutputTimestampFormat);
                    break;
                case eOutputFormat.TABLE:
                    OutputTable(writer, OPCHDAItemValues, OutputTimestampFormat);
                    break;
            }

            if (!string.IsNullOrEmpty(OutputFileName)) {
                writer.Close();
                Console.WriteLine("Data were written to file {0}.", OutputFileName);
            }
            
            return;
        }

        static void OutputList(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            string ts;
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                sw.WriteLine();
                sw.WriteLine("\tTag ({0} of {1}): {2} ({3} values):", i + 1, OPCHDAItemValues.Count(), OPCHDAItemValues[i].ItemName, OPCHDAItemValues[i].Count);
                sw.WriteLine("{0,20}{1,20}{2,20}", "Timestamp", "Value", "Quality");

                for (int j = 0; j < OPCHDAItemValues[i].Count; j++) {
                    if (OutputTimestampFormat != null) {
                        ts = OPCHDAItemValues[i][j].Timestamp.ToString();
                    } else {
                        ts = OPCHDAItemValues[i][j].Timestamp.ToString(OutputTimestampFormat);
                    }
                    sw.WriteLine("{0,20}{1,20}{2,20}", ts, OPCHDAItemValues[i][j].Value.ToString(), OPCHDAItemValues[i][j].Quality.ToString());
                }
            }
        }
        static void OutputTable(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            // header
            sw.Write("{0} timestamp, {0} value, {0} quality", OPCHDAItemValues[0].ItemName);
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                sw.Write(",{0} timestamp, {0} value, {0} quality", OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();
            string ts;
            // What if different tags have different number of points?!
            for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                ts = GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, OutputTimestampFormat);
                sw.Write("{0},{1},{2}", ts, OPCHDAItemValues[0][j].Value.ToString(), OPCHDAItemValues[0][j].Quality.ToString());
                for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                    ts = GetDatetimeStr(OPCHDAItemValues[i][j].Timestamp, OutputTimestampFormat);
                    sw.Write(",{0},{1},{2}", ts, OPCHDAItemValues[i][j].Value.ToString(), OPCHDAItemValues[i][j].Quality.ToString());
                }
                sw.WriteLine();
            }
        }
        static void OutputMerged(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            // header
            sw.Write("Timestamp, {0} value, {0} quality", OPCHDAItemValues[0].ItemName);
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                sw.Write(",{0} value, {0} quality", OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();
            string ts;
            // What if different tags have different number of points?!
            for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                ts = GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, OutputTimestampFormat);
                sw.Write("{0},{1},{2}", ts, OPCHDAItemValues[0][j].Value.ToString(), OPCHDAItemValues[0][j].Quality.ToString());
                for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                    if (OPCHDAItemValues[i][j].Timestamp != OPCHDAItemValues[0][j].Timestamp) {
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Timestamp {0} of tag {1} at row {2} is not equal to timestamp {3} of tag {4}",
                            OPCHDAItemValues[i][j].Timestamp, OPCHDAItemValues[i].ItemName, j, OPCHDAItemValues[0][j].Timestamp, OPCHDAItemValues[0].ItemName);
                    }
                    sw.Write(",{0},{1}", OPCHDAItemValues[i][j].Value.ToString(), OPCHDAItemValues[i][j].Quality.ToString());
                }
                sw.WriteLine();
            }
        }
        static void ShowHelp(OptionSet p) {
            Console.WriteLine("Usage: HDARead OPTIONS tag1 tag2 tag3");
            Console.WriteLine("HDARead is used to read the data from OPC HDA server.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
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
                throw new NDesk.Options.OptionException("Wrong output format:: " + str, "-f");
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
