using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NDesk.Options;
using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

namespace HDARead {
    class Program {

        static TraceSource _trace = new TraceSource("ConsoleApplicationTraceSource");

        static void Main(string[] args) {

            // Private NumberOfTags As ConInt
            string Host = null, Server = null;
            string StartTime = null, EndTime = null;
            int Aggregate = 3; // Average
            string OutputTimestampFormat = "";
            int MaxValues = 10;
            int ResampleInterval = 0;
            bool show_help = false;
            List<string> Tagnames = new List<string>();

            var p = new OptionSet() {
   	            { "n=|node=",               v => Host = v },
   	            { "s=|server=",             v => Server = v },
   	            { "from=|start=|begin=",    v => StartTime = v ?? "NOW-1H"},
   	            { "to=|end=",               v => EndTime = v ?? "NOW"},
   	            { "a=|agg=",                v => Aggregate = Int32.Parse(v) },
                { "f=|tsformat=",           v => OutputTimestampFormat = v ?? ""},
                { "m=|maxvalues=",          v => MaxValues = Int32.Parse(v)},
                { "r=|resample=",           v => ResampleInterval = Int32.Parse(v)},
   	            { "h|?|help",               v => show_help = v != null },
                { "<>","List of tag names", v => Tagnames.Add (v)},
            };

            List<string> extra;
            try {
                extra = p.Parse(args);
            } catch (OptionException e) {
                Console.Write("HDARead: ");
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

            //string[] ArrResults = new string[Tagnames.Count()];
            //string[] ArrTimestamps = new string[Tagnames.Count()];
            //string[] ArrQualities = new string[Tagnames.Count()];
            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;

            try {
                bool res;
                var srv = new HDAClient();
                _trace.TraceEvent(TraceEventType.Information, 0, "Created HDAClient");
                if (srv.Connect(Host, Server)) {
                    _trace.TraceEvent(TraceEventType.Information, 0, "Connected. Going to read.");
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, out OPCHDAItemValues);
                    if (res) {
                        _trace.TraceEvent(TraceEventType.Information, 0, "HDARead OK.");
                        Console.WriteLine("HDARead OK.");
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "OPCHDAItemValues.Count()={0}", OPCHDAItemValues.Count());
                        for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                            _trace.TraceEvent(TraceEventType.Verbose, 0, "OPCHDAItemValues[i].Count={0}", OPCHDAItemValues[i].Count);
                            for (int j = 0; j < OPCHDAItemValues[i].Count; j++) {
                                if (OutputTimestampFormat == "") {
                                    Console.Write(OPCHDAItemValues[i][j].Timestamp.ToString());
                                } else {
                                    Console.Write(OPCHDAItemValues[i][j].Timestamp.ToString(OutputTimestampFormat));
                                }
                                Console.Write(", v=" + OPCHDAItemValues[i][j].Value.ToString() + ",");
                                Console.WriteLine();
                            }
                        }

                        /*
                        //Debug.Print(OPCHDAItemValues(0).ItemName, OPCHDAItemValues(0).ResultID.Name)
                        if ((OPCHDAItemValues[0][0].Value != null) && (OPCHDAItemValues[0].ResultID == Opc.ResultID.S_OK)) {
                            ArrResults[i] = OPCHDAItemValues[0][0].Value.ToString();
                            //Convert timestamp according to OutputTimestampFormat
                            if (OutputTimestampFormat == "") {
                                ArrTimestamps[i] = OPCHDAItemValues[0][0].Timestamp.ToString();
                            } else {
                                ArrTimestamps[i] = OPCHDAItemValues[0][0].Timestamp.ToString(OutputTimestampFormat);
                            }
                            ArrQualities[i] = OPCHDAItemValues[0][0].Quality.ToString();
                        } else {
                            //TraceS("History2: item " & i & " " & ArrTagnames(i) & ": ResultID=" & OPCHDAItemValues(0).ResultID.ToString)
                            ArrResults[i] = "NaN";
                            ArrTimestamps[i] = "";
                            ArrQualities[i] = "ERROR";
                        }
                        */

                    } else {
                        Console.WriteLine("HDARead Error.");
                    }
                }
                srv.Disconnect();
            } catch (Exception e) {
                Console.WriteLine(e.Message);
            }
            return;
        }

        static void ShowHelp(OptionSet p) {
            Console.WriteLine("Usage: HDARead OPTIONS tag1 tag2 tag3");
            Console.WriteLine("HDARead is used to read the data from OPC HDA server.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
