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
 * + remember that timestamps may be in reverse order. Currently this will crash 'Merge' algorithm and may be something else
 * partition the program to separate modules/classes...
 * write to file portion by portion to conserve memory...
 * check and rewise exceptions vs. error codes
 * better merge algorithm?
 * if output file already exists: overwrite or do nothing
 * omit value if NODATA quality
 * shutdown event
 * log to file?
 * * */

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

            try {
                var out_writer = new OutputWriter(OutputFormat, OutputQuality, OutputFileName, OutputTimestampFormat, ReadRaw, _trace.Switch.Level);
                out_writer.WriteHeader(OPCHDAItemValues);
                out_writer.Write(OPCHDAItemValues);
                out_writer.Close();
            } catch (Exception e) {
                Console.WriteLine(e.Message);
                return;
            }
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
                { "m=|maxvalues=",          "Maximum number of values to load (only for raw data)", 
                                                                                v => MaxValues = Int32.Parse(v)},
                { "b|bounds",               "Whether the bounding item values should be returned (only for ReadRaw).",  
                                                                                v => IncludeBounds = v != null},
                { "t=|tsformat=",           "Output timestamp format to use",   v => OutputTimestampFormat = v},
                { "f=",                     "Output format (TABLE or MERGED)",   
                                                                                v => OutputFormat = Utils.GetOutputFormat(v)},
                { "q=",                     "Include quality in output data (NONE, DA, HISTORIAN or BOTH)",   
                                                                                v => OutputQuality = Utils.GetOutputQuality(v)},
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
