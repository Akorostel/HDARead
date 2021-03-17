using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Globalization;

namespace HDARead {
    abstract public class OutputWriter {

        static protected TraceSource _trace = new TraceSource("OutputWriterTraceSource");
        protected StreamWriter _writer = null;
        protected eOutputFormat _OutputFormat = eOutputFormat.MERGED;
        protected eOutputQuality _OutputQuality = eOutputQuality.NONE;
        protected string _OutputFileName = null;
        protected string _OutputTimestampFormat = null;
        protected bool _ReadRaw = false;

        public OutputWriter(eOutputFormat OutputFormat,
                            eOutputQuality OutputQuality,
                            string OutputFileName,
                            string OutputTimestampFormat,
                            bool ReadRaw,
                            SourceLevels swlvl = SourceLevels.Information) {

            _trace.Switch.Level = swlvl;
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Creating OutputWriter.");

            _OutputFormat = OutputFormat;
            _OutputQuality = OutputQuality;
            _OutputFileName = OutputFileName;
            _OutputTimestampFormat = OutputTimestampFormat;
            _ReadRaw = ReadRaw;

            try {
                if (!string.IsNullOrEmpty(_OutputFileName)) {
                    _writer = new StreamWriter(_OutputFileName);
                } else {
                    _writer = new StreamWriter(Console.OpenStandardOutput());
                    _writer.AutoFlush = true;
                    Console.SetOut(_writer);
                }

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during creating OutputWriter:" + e.ToString());
                Close();
                throw;
            }
        }

        public void Close() {
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Closing OutputWriter.");
            if (!string.IsNullOrEmpty(_OutputFileName) && (_writer != null)) {
                _writer.Close();
            }
        }

        abstract public void WriteHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues);
        abstract public void Write(Opc.Hda.ItemValueCollection[] OPCHDAItemValues);
    }


    // Output format: merged
    // timestamp, tag1 value, tag2 value, ...
    public class MergedOutputWriter : OutputWriter {
        public MergedOutputWriter(eOutputFormat OutputFormat,
                    eOutputQuality OutputQuality,
                    string OutputFileName,
                    string OutputTimestampFormat,
                    bool ReadRaw,
                    SourceLevels swlvl = SourceLevels.Information) : base(OutputFormat, 
                                                                          OutputQuality, 
                                                                          OutputFileName, 
                                                                          OutputTimestampFormat, 
                                                                          ReadRaw, 
                                                                          swlvl)
        { }

        public override void WriteHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                if (_OutputTimestampFormat == "DateTime")
                    _writer.Write("Date,Time");
                else
                    _writer.Write("Timestamp");

                string hdr = ",{0}";
                if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                    hdr += ",{0} da quality";
                }
                if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                    hdr += ",{0} hist quality";
                }
                // header
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    _writer.Write(hdr, OPCHDAItemValues[i].ItemName);
                }
                _writer.WriteLine();
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output header:" + e.ToString());
                Close();
                throw;
            }
        }

        public override void Write(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                if (_ReadRaw) {
                    Merger.SetDebugLevel(_trace.Switch.Level);
                    OPCHDAItemValues = Merger.Merge(OPCHDAItemValues);
                }

                for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                    if (_OutputTimestampFormat == "DateTime")
                        _writer.Write("{0},{1}",
                            OPCHDAItemValues[0][j].Timestamp.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture),
                            OPCHDAItemValues[0][j].Timestamp.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                    else
                        _writer.Write("{0}", Utils.GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, _OutputTimestampFormat));
                    for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                        // Maybe its better to catch exception (null ref) than to check every element
                        if (OPCHDAItemValues[i][j].Value == null) {
                            _writer.Write(", ");
                        } else {
                            _writer.Write(",{0}", OPCHDAItemValues[i][j].Value.ToString());
                        }

                        if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                            if (OPCHDAItemValues[i][j].Quality == null) {
                                _writer.Write(",");
                            } else {
                                _writer.Write(",{0}", OPCHDAItemValues[i][j].Quality.ToString());
                            }
                        }

                        if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                            // OPC.DA.Quality is struct, but OPC.HDA.Quality is enum.
                            // Enum cannot be null, so there is no need to check
                            _writer.Write(",{0}", OPCHDAItemValues[i][j].HistorianQuality.ToString());
                        }
                    }
                    _writer.WriteLine();
                }

                if (!string.IsNullOrEmpty(_OutputFileName)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Data were written to file {0}.", _OutputFileName);
                }
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output:" + e.ToString());
                Close();
                throw;
            } 
        }
    }

    // Output format: table
    // tag1 timestamp, tag1 value, tag2 timestamp, tag2 value, ...
    public class TableOutputWriter : OutputWriter {
        public TableOutputWriter(eOutputFormat OutputFormat,
                    eOutputQuality OutputQuality,
                    string OutputFileName,
                    string OutputTimestampFormat,
                    bool ReadRaw,
                    SourceLevels swlvl = SourceLevels.Information) : base(OutputFormat, 
                                                                          OutputQuality, 
                                                                          OutputFileName, 
                                                                          OutputTimestampFormat, 
                                                                          ReadRaw, 
                                                                          swlvl)
        { }

        public override void WriteHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                string hdr;
                if (_OutputTimestampFormat == "DateTime")
                    hdr = "Date,Time,{0}";
                else
                    hdr = "Timestamp,{0}";

                if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                    hdr += ", {0} da quality";
                }
                if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                    hdr += ", {0} hist quality";
                }

                _writer.Write(hdr, OPCHDAItemValues[0].ItemName);
                for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                    _writer.Write(",");
                    _writer.Write(hdr, OPCHDAItemValues[i].ItemName);
                }
                _writer.WriteLine();
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output header:" + e.ToString());
                Close();
                throw;
            }
        }

        public override void Write(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            try {
                string valstr = ",{0}";
                string emptystr;
                if (_OutputTimestampFormat == "DateTime")
                    emptystr = ",,";
                else
                    emptystr = ",";

                if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH)) {
                    valstr += ",{1}";
                    emptystr += ",";
                }
                if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH)) {
                    valstr += ",{2}";
                    emptystr += ",";
                }

                // What if different tags have different number of points?!
                // This shouldn't be possible.
                int max_rows = OPCHDAItemValues.Max(x => x.Count);

                for (int j = 0; j < max_rows; j++) {
                    for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                        if (i > 0) {
                            _writer.Write(",");
                        }

                        if (j < OPCHDAItemValues[i].Count) {
                            if (_OutputTimestampFormat == "DateTime")
                                _writer.Write("{0},{1}",
                                    OPCHDAItemValues[0][j].Timestamp.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture),
                                    OPCHDAItemValues[0][j].Timestamp.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                            else
                                _writer.Write("{0}", Utils.GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, _OutputTimestampFormat));

                            _writer.Write(valstr,
                                ((OPCHDAItemValues[i][j].Value == null) ? "" : OPCHDAItemValues[i][j].Value.ToString()),
                                OPCHDAItemValues[i][j].Quality.ToString(),
                                OPCHDAItemValues[i][j].HistorianQuality.ToString());
                        } else {
                            _writer.Write(emptystr);
                        }
                    }
                    _writer.WriteLine();
                }

                if (!string.IsNullOrEmpty(_OutputFileName)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Data were written to file {0}.", _OutputFileName);
                }
                return;

            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output:" + e.ToString());
                Close();
                throw;
            }
        }
    }

    // Output format: record
	// timestamp1, tag1, ts1_tag1_value
	// timestamp1, tag2, ts1_tag2_value
	// timestamp2, tag1, ts2_tag1_value
	// timestamp3, tag2, ts3_tag2_value
    // ...
    public class RecordOutputWriter : OutputWriter
    {
        public RecordOutputWriter(eOutputFormat OutputFormat,
                    eOutputQuality OutputQuality,
                    string OutputFileName,
                    string OutputTimestampFormat,
                    bool ReadRaw,
                    SourceLevels swlvl = SourceLevels.Information) : base(OutputFormat,
                                                                          OutputQuality,
                                                                          OutputFileName,
                                                                          OutputTimestampFormat,
                                                                          ReadRaw,
                                                                          swlvl)
        { }

        public override void Write(Opc.Hda.ItemValueCollection[] OPCHDAItemValues)
        {
            try
            {
                string valstr = ",{0}";
                if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH))
                    valstr += ",{1}";
                if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH))
                    valstr += ",{2}";
                foreach (Opc.Hda.ItemValueCollection tagValues in OPCHDAItemValues)
                {
                    foreach (Opc.Hda.ItemValue tagValue in tagValues)
                    {
                        if (_OutputTimestampFormat == "DateTime")
                            _writer.Write("{0},{1}",
                                tagValue.Timestamp.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture),
                                tagValue.Timestamp.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                        else
                            _writer.Write("{0}", Utils.GetDatetimeStr(tagValue.Timestamp, _OutputTimestampFormat));
                        _writer.Write(",{0}", tagValues.ItemName);
                        _writer.Write(valstr,
                            ((tagValue.Value == null) ? "" : tagValue.Value.ToString()),
                            tagValue.Quality.ToString(),
                            tagValue.HistorianQuality.ToString());
                        _writer.WriteLine();
                    }
                }
                if (!string.IsNullOrEmpty(_OutputFileName))
                {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Data were written to file {0}.", _OutputFileName);
                }
            }
            catch (Exception e)
            {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output:" + e.ToString());
                Close();
                throw;
            }
        }

        public override void WriteHeader(Opc.Hda.ItemValueCollection[] OPCHDAItemValues)
        {
            try
            {
                string hdr;
                if (_OutputTimestampFormat == "DateTime")
                    hdr = "Date,Time";
                else
                    hdr = "Timestamp";
                hdr += ",Tag,Value";
                if ((_OutputQuality == eOutputQuality.DA) || (_OutputQuality == eOutputQuality.BOTH))
                {
                    hdr += ",DA quality";
                }
                if ((_OutputQuality == eOutputQuality.HISTORIAN) || (_OutputQuality == eOutputQuality.BOTH))
                {
                    hdr += ",Hist quality";
                }
                _writer.Write(hdr);
                _writer.WriteLine();
                return;

            }
            catch (Exception e)
            {
                _trace.TraceEvent(TraceEventType.Error, 0, "Exception during writing output header:" + e.ToString());
                Close();
                throw;
            }
        }
    }
}