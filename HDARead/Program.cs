using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NDesk.Options;

namespace HDARead {
    class Program {
        static void Main(string[] args) {

            // Private NumberOfTags As ConInt
            string Host = null, Server = null;
            string StartTime = null, EndTime = null;
            string Aggregate;
            string OutputTimestampFormat;
            bool show_help = false;
            List<string> Tagnames = new List<string>();

            var p = new OptionSet() {
   	            { "n=|node=",               v => Host = v },
   	            { "s=|server=",             v => Server = v },
   	            { "from=|start=|begin=",    v => StartTime = v ?? "NOW-1H"},
   	            { "to=|end=",               v => EndTime = v ?? "NOW"},
   	            { "a=|agg=",                v => Aggregate = v },
                { "f=|tsformat=",           v => OutputTimestampFormat = v },
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
        }

        static void ShowHelp(OptionSet p) {
            Console.WriteLine("Usage: HDARead [OPTIONS]+");
            Console.WriteLine("HDARead is used to read the data from OPC HDA server.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
