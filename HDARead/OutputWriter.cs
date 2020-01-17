using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

using System.IO;

namespace HDARead {
    class OutputWriter {

        static TraceSource _trace = new TraceSource("OutputWriterTraceSource");
        StreamWriter writer = null;
        eOutputFormat _OutputFormat = eOutputFormat.MERGED;
        eOutputQuality _OutputQuality = eOutputQuality.NONE;
        string _OutputFileName = null;
        string _OutputTimestampFormat = null;
        bool _ReadRaw = false;

        public OutputWriter(eOutputFormat OutputFormat,
                            eOutputQuality OutputQuality, 
                            string OutputFileName, 
                            string OutputTimestampFormat, 
                            bool ReadRaw,
                            SourceLevels swlvl = SourceLevels.Information) {
            _trace = new TraceSource("OutputWriterTraceSource");
            _trace.Switch.Level = swlvl;
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Creating OutputWriter.");

            _OutputFormat = OutputFormat;
            _OutputQuality = OutputQuality;
            _OutputFileName = OutputFileName;
            _OutputTimestampFormat = OutputTimestampFormat;
            _ReadRaw = ReadRaw;

            try {
                if (!string.IsNullOrEmpty(_OutputFileName)) {
                    writer = new StreamWriter(_OutputFileName);
                } else {
                    writer = new StreamWriter(Console.OpenStandardOutput());
                    writer.AutoFlush = true;
                    Console.SetOut(writer);
                }

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during creating OutoutWriter:" + e.ToString());
                if (!string.IsNullOrEmpty(_OutputFileName) && (writer != null)) {
                    writer.Close();
                }
                throw;
            } 
        }

        public void Close() {
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Closing OutputWriter.");
            if (!string.IsNullOrEmpty(_OutputFileName) && (writer != null)) {
                writer.Close();
            }
        }

        public void WriteHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                switch (_OutputFormat) {
                    case eOutputFormat.MERGED:
                        OutputMergedHeader(OPCHDAItemValues);
                        break;
                    case eOutputFormat.TABLE:
                        OutputTableHeader(OPCHDAItemValues);
                        break;
                }
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output header:" + e.ToString());
                if (!string.IsNullOrEmpty(_OutputFileName) && (writer != null)) {
                    writer.Close();
                }
                throw;
            } 
        }

        public void Write(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                switch (_OutputFormat) {
                    case eOutputFormat.MERGED:
                        if (_ReadRaw) {
                            var MyMerger = new Merger(_trace.Switch.Level);
                            OutputMerged(MyMerger.Merge(OPCHDAItemValues));
                        } else {
                            OutputMerged(OPCHDAItemValues);
                        }
                        break;
                    case eOutputFormat.TABLE:
                        OutputTable(OPCHDAItemValues);
                        break;
                }

                if (!string.IsNullOrEmpty(_OutputFileName)) {
                    Console.WriteLine("Data were written to file {0}.", _OutputFileName);
                }
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output:" + e.ToString());
                throw;
            } finally {
                if (!string.IsNullOrEmpty(_OutputFileName) && (writer != null)) {
                    writer.Close();
                }
            }
        }

        // Output format: merged
        // timestamp, tag1 value, tag2 value, ...
        void OutputMergedHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            writer.Write("Timestamp");
            string hdr = ", {0}";
            if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} da quality";
            }
            if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} hist quality";
            }
            // header
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                writer.Write(hdr, OPCHDAItemValues[i].ItemName);
            }
            writer.WriteLine();
        }

        void OutputMerged(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                writer.Write("{0}", Utils.GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, _OutputTimestampFormat));
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    // Maybe its better to catch exception (null ref) than to check every element
                    if (OPCHDAItemValues[i][j].Value == null) {
                        writer.Write(", ");
                    } else {
                        writer.Write(", {0}", OPCHDAItemValues[i][j].Value.ToString());
                    }

                    if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                        if (OPCHDAItemValues[i][j].Quality == null) {
                            writer.Write(", ");
                        } else {
                            writer.Write(", {0}", OPCHDAItemValues[i][j].Quality.ToString());
                        }
                    }

                    if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                        // OPC.DA.Quality is struct, but OPC.HDA.Quality is enum.
                        // Enum cannot be null, so there is no need to check
                        writer.Write(", {0}", OPCHDAItemValues[i][j].HistorianQuality.ToString());
                    }
                }
                writer.WriteLine();
            }
        }

        // Output format: table
        // tag1 timestamp, tag1 value, tag2 timestamp, tag2 value, ...
        void OutputTableHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            string hdr = "Timestamp, {0}";

            if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} da quality";
            }
            if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                hdr += ", {0} hist quality";
            }
            // header
            writer.Write(hdr, OPCHDAItemValues[0].ItemName);
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                writer.Write(", ");
                writer.Write(hdr, OPCHDAItemValues[i].ItemName);
            }
            writer.WriteLine();
        }

        // Output format: table
        // tag1 timestamp, tag1 value, tag2 timestamp, tag2 value, ...
        void OutputTable(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            string valstr = "{0},{1}";
            string emptystr = ",";

            if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                valstr += ",{2}";
                emptystr += ",";
            }
            if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                valstr += ",{3}";
                emptystr += ",";
            }

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

            for (int j = 0; j < max_rows; j++) {
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    if (i > 0) {
                        writer.Write(", ");
                    }
                    if (j < OPCHDAItemValues[i].Count) {
                        writer.Write(valstr,
                            Utils.GetDatetimeStr(OPCHDAItemValues[i][j].Timestamp, _OutputTimestampFormat),
                            OPCHDAItemValues[i][j].Value.ToString(),
                            OPCHDAItemValues[i][j].Quality.ToString(),
                            OPCHDAItemValues[i][j].HistorianQuality.ToString());
                    } else {
                        writer.Write(emptystr);
                    }
                }
                writer.WriteLine();
            }
        }
    }
}
