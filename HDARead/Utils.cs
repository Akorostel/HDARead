using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Opc;
using OpcCom;

namespace HDARead {
    static class Utils {

        static public int GetHDAAggregate(string str) {
            HDAClient.OPCHDA_AGGREGATE Value;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(HDAClient.OPCHDA_AGGREGATE), Value))
                return (int)Value;
            else
                throw new NDesk.Options.OptionException("Wrong aggregate: " + str, "-a");
        }

        static public eOutputFormat GetOutputFormat(string str) {
            eOutputFormat Value;
            if (string.IsNullOrEmpty(str))
                return eOutputFormat.MERGED;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(eOutputFormat), Value))
                return Value;
            else
                throw new NDesk.Options.OptionException("Wrong output format: " + str, "-f");
        }

        static public eOutputQuality GetOutputQuality(string str) {
            eOutputQuality Value;
            if (string.IsNullOrEmpty(str))
                return eOutputQuality.NONE;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(eOutputQuality), Value))
                return Value;
            else
                throw new NDesk.Options.OptionException("Wrong output quality: " + str, "-q");
        }

        static public string GetDatetimeStr(DateTime dt, string fmt) {
            string ts;
            if (string.IsNullOrEmpty(fmt)) {
                ts = dt.ToString();
            } else {
                ts = dt.ToString(fmt);
            }
            return ts;
        }

        static public void ConsoleWriteColoredLine(ConsoleColor color, String value) {
            var originalForegroundColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(value);
            Console.ForegroundColor = originalForegroundColor;
        }

        static public void ListHDAServers(String node) {
            IDiscovery discovery = new OpcCom.ServerEnumerator();
            Opc.Server[] servers;

            if (string.IsNullOrEmpty(node)) 
                servers = discovery.GetAvailableServers(Specification.COM_HDA_10);
            else
                servers = discovery.GetAvailableServers(Specification.COM_HDA_10, node, null);

            foreach (Opc.Server s in servers) {
                Console.WriteLine(s.Name);
            }
        }

    }
}