using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NDesk.Options;
using System.Diagnostics;
using System.IO;
using System.Reflection;

/* TODO:
 * better merge algorithm?
 * if output file already exists: overwrite or do nothing
 * specify decimal point symbol and value separator
 * check if program works with string tags
 * omit value if NODATA quality
 * empty column if tag doesnt exist
 * handle shutdown event
 * log to file?
 * use job configuration file instead of specifying all parameters in command line
 * * */

namespace HDARead {

    public enum eOutputFormat {
        LIST = 1,
        TABLE = 2,
        MERGED = 3,
        RECORD = 4
    }
    public enum eOutputQuality {
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
        static bool ExtendedInfo = false;
        static eOutputQuality OutputQuality = eOutputQuality.NONE;
        static eOutputFormat OutputFormat = eOutputFormat.MERGED;
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
            if (ExtendedInfo) {
                Console.WriteLine("HDARead v. " + Assembly.GetExecutingAssembly().GetName().Version.ToString());
                ShowInfo();
            }
            var srv = new HDAClient(_trace.Switch);

            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;
            try {
                bool res = false;
                if (srv.Connect(Host, Server)) {
                    // Remove unknown items from the list
                    srv.Validate(Tagnames);
                    // Create OutputWriter
                    OutputWriter out_writer;
                    switch (OutputFormat) {
                        case eOutputFormat.MERGED:
                            out_writer = new MergedOutputWriter(OutputFormat, OutputQuality, OutputFileName, OutputTimestampFormat, ReadRaw, _trace.Switch.Level);
                            break;
                        case eOutputFormat.TABLE:
                            out_writer = new TableOutputWriter(OutputFormat, OutputQuality, OutputFileName, OutputTimestampFormat, ReadRaw, _trace.Switch.Level);
                            break;
                        case eOutputFormat.RECORD:
                            out_writer = new RecordOutputWriter(OutputFormat, OutputQuality, OutputFileName, OutputTimestampFormat, ReadRaw, _trace.Switch.Level);
                            break;
                        default:
                            throw (new ArgumentException("Unknown output format"));
                    }
                    // Read items 
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, IncludeBounds, ReadRaw, out_writer, out OPCHDAItemValues);
                } else {
                    Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "HDARead unable to connect to OPC server.");
                }
                if (!res) {
                    Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "Error reading data.");
                    return;
                }
            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, e.Message);
                Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "Error reading data.");
                return;
            } finally {
                srv.Disconnect();
            }
            Utils.ConsoleWriteColoredLine(ConsoleColor.Green, "Data read OK.");
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
   	            { "a=|agg=",                "Aggregate (see spec)",             v => Aggregate = Utils.GetHDAAggregate(v) },
                { "r=|resample=",           "Resample interval (in seconds), 0 - return just one value (see OPC HDA spec.)",  
                                                                                v => ResampleInterval = Int32.Parse(v)},
                { "raw",                    "Read raw data (if omitted, read processed data) ",  
                                                                                v => ReadRaw = v != null},
                { "m=|maxvalues=",          "Maximum number of values to load (only for ReadRaw)", 
                                                                                v => MaxValues = Int32.Parse(v)},
                { "b|bounds",               "Whether the bounding item values should be returned (only for ReadRaw).",  
                                                                                v => IncludeBounds = v != null},
                { "t=|tsformat=",           "Output timestamp format to use. You can use -t=DateTime to output date and time in separate columns",  
                                                                                v => OutputTimestampFormat = v},
                { "f=",                     "Output format (TABLE, MERGED or RECORD)",   
                                                                                v => OutputFormat = Utils.GetOutputFormat(v)},
                { "q=",                     "Include quality in output data (NONE, DA, HISTORIAN or BOTH)",   
                                                                                v => OutputQuality = Utils.GetOutputQuality(v)},
                { "o=|output=",             "Output filename (if omitted, output to console)",   
                                                                                v => OutputFileName = v},
                { "i=|input=",              "Input filename with list of tags (if omitted, tag list must be provided as command line argument)",   
                                                                                v => InputFileName = v},
                { "v",                      "Show extended info",               v => ExtendedInfo = v != null},
                { "vv|verbose",             "Show debug info",                  v => Verbose = v != null},
                { "h|?|help",               "Show help",                        v => Help = v != null},
                { "<>",                     "List of tag names",                v => Tagnames.Add (v)},
            };

            try {
                p.Parse(args);
            } catch (OptionException e) {
                Utils.ConsoleWriteColoredLine(ConsoleColor.Red, e.Message);
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
                Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "Missing required option s=|server=");
                Console.WriteLine("Available servers are:");
                Utils.ListHDAServers(Host);
                return false;
            }

            if (string.IsNullOrEmpty(InputFileName)) {
                if (Tagnames.Count() < 1) {
                    Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "No tagnames were specified.");
                    return false;
                }
            } else {
                if (Tagnames.Count() > 0) {
                    Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "If the input file is specified, no tags may be entered as command line argument");
                    return false;
                }
                try {
                    Tagnames = File.ReadLines(InputFileName).ToList();
                    if (Tagnames.Count() < 1) {
                        Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "No tagnames were specified.");
                        return false;
                    }
                } catch (Exception e) {
                    _trace.TraceEvent(TraceEventType.Error, 0, e.Message);
                    Utils.ConsoleWriteColoredLine(ConsoleColor.Red, "Error reading tags from file: " + e.Message);
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

            if (Verbose || Tagnames.Count<10) {
                foreach (string t in Tagnames) {
                    Console.WriteLine("\t\t" + t);
                }
            } else {
                int n = Tagnames.Count;
                for (int i = 0; i < 3; i++) {
                    Console.WriteLine("\t\t" + Tagnames[i]);
                }
                Console.WriteLine("\t\t..." );
                for (int i = n-3; i < n; i++) {
                    Console.WriteLine("\t\t" + Tagnames[i]);
                }
            }

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

    }
}