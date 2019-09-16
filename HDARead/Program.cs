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

/* TODO:
 * aggregate enum
 * resample interval
 * file output
 * quality flag
 * exceptions
 * read raw
 * input file with tag list
 * maybe its better to query data tag by tag (because 'tag not found' exception doesn't show which tag is wrong)
 */

namespace HDARead {
    class Program {

        static TraceSource _trace = new TraceSource("ConsoleApplicationTraceSource");

        static void Main(string[] args) {

            // Private NumberOfTags As ConInt
            string Host = null, Server = null;
            string StartTime = null, EndTime = null;
            int Aggregate = (int)HDAClient.OPCHDA_AGGREGATE.AVERAGE;
            string OutputTimestampFormat = "";
            int MaxValues = 10;
            int ResampleInterval = 0;
            bool show_help = false;
            List<string> Tagnames = new List<string>();

            var p = new OptionSet() {
   	            { "n=|node=",               "Remote computer",                  v => Host = v },
   	            { "s=|server=",             "OPC HDA server name (required)",   v => Server = v },
   	            { "from=|start=|begin=",    "Start time (abs. or relative)",    v => StartTime = v ?? "NOW-1H"},
   	            { "to=|end=",               "End time (abs. or relative)",      v => EndTime = v ?? "NOW"},
   	            { "a=|agg=",                "Aggregate (see spec)",             v => Aggregate = GetHDAAggregate(v) },
                { "r=|resample=",           "Resample interval (in seconds), 0 - return just one value (see OPC HDA spec.)",  
                                                                                v => ResampleInterval = Int32.Parse(v)},
                { "m=|maxvalues=",          v => MaxValues = Int32.Parse(v)},
                { "f=|tsformat=",           v => OutputTimestampFormat = v ?? ""},
                { "h|?|help",               v => show_help = v != null },
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

            if (Server == null) {
                Console.WriteLine("Missing required option s=|server=");
                return;
            }

            if (Tagnames.Count()<1) {
                Console.WriteLine("No tagnames were specified.");
                return;
            }

            //Default values
            StartTime = StartTime ?? "NOW-1H";
            EndTime = EndTime ?? "NOW";

            Console.WriteLine("HDARead is going to:");
            Console.WriteLine("\t connect to OPC HDA server named '{0}' on computer '{1}'", Server, Host);
            Console.WriteLine("\t get data for the period from '{0}' to '{1}'", StartTime, EndTime);
            Console.WriteLine("\t for the following tags:");
            foreach (string t in Tagnames) {
                Console.WriteLine("\t\t" + t );
            }

            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;

            try {
                bool res=false;
                var srv = new HDAClient();
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Created HDAClient");
                if (srv.Connect(Host, Server)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Connected. Going to read.");
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, out OPCHDAItemValues);
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
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Number of data points = OPCHDAItemValues[i].Count={0}", OPCHDAItemValues[i].Count);
                for (int j = 0; j < OPCHDAItemValues[i].Count; j++) {
                    if (OutputTimestampFormat == "") {
                        Console.Write(OPCHDAItemValues[i][j].Timestamp.ToString());
                    } else {
                        Console.Write(OPCHDAItemValues[i][j].Timestamp.ToString(OutputTimestampFormat));
                    }
                    Console.Write(", " + OPCHDAItemValues[i][j].Value.ToString());
                    Console.Write(", " + OPCHDAItemValues[i][j].Quality.ToString());
                    Console.WriteLine();
                }
            }

            return;
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
                Console.WriteLine("\t" + agg);
            }
                
        }

        static int GetHDAAggregate(string str) {
            HDAClient.OPCHDA_AGGREGATE Value;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(HDAClient.OPCHDA_AGGREGATE), Value))
                return (int)Value;
            else
                throw new NDesk.Options.OptionException("Wrong aggregate: " + str, "-a");
        }
    }
}
