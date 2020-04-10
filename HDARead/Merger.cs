using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

namespace HDARead {
    static class Merger {
        static TraceSource _trace = new TraceSource("MergerTraceSource");

        public static void SetDebugLevel(SourceLevels swlvl = SourceLevels.Information) {
            _trace.Switch.Level = swlvl;
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Created Merger");
        }

        // Merge multiple timeseries. Fill with NaN.
        public static Opc.Hda.ItemValueCollection[] Merge(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
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
    }
}
