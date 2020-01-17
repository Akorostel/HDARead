using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Opc;
using OpcCom;

using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

namespace HDARead {
    class HDAClient {

        private Opc.Hda.Server _OPCServer = null;
        public Opc.Hda.ServerStatus Status = null;
        public Opc.Hda.Aggregate[] SupportedAggregates = null;
        private TraceSource _trace; 

        public enum OPCHDA_AGGREGATE {
            ANNOTATIONS = 24,
            WORSTQUALITY = 23,
            PERCENTBAD = 22,
            PERCENTGOOD = 21,
            DURATIONBAD = 20,
            DURATIONGOOD = 19,
            RANGE = 18,
            VARIANCE = 17,
            REGDEV = 16,
            REGCONST = 15,
            REGSLOPE = 14,
            DELTA = 13,
            END = 12,
            START = 11,
            MAXIMUM = 10,
            MAXIMUMACTUALTIME = 9,
            MINIMUM = 8,
            MINIMUMACTUALTIME = 7,
            STDEV = 6,
            COUNT = 5,
            TIMEAVERAGE = 4,
            AVERAGE = 3,
            TOTAL = 2,
            INTERPOLATIVE = 1,
            NOAGGREGATE = 0
        }

        public HDAClient() {
            _trace = new TraceSource("HDAClientTraceSource");
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Created HDAClient");
        }

        // This constructor sets verbosity level to HDACLient instance
        public HDAClient(SourceSwitch sw) {
            _trace = new TraceSource("HDAClientTraceSource");
            _trace.Switch = sw;
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Created HDAClient");
        }

        // Connect to OPC HDA server
        public bool Connect(string HostName, string ServerName) {

            var url = new Opc.URL();

            url.Scheme = Opc.UrlScheme.HDA;
            url.HostName = HostName;
            url.Path = ServerName;

            try {
                var fact = new OpcCom.Factory();
                if ((_OPCServer != null) && (!_OPCServer.IsConnected)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "_OPCServer is disconnected, disposing object");
                    //Unfortunately, in case of lost connection simply calling .Connect() doesn't work :(
                    //Let's try to recreate the object from scratch
                    _OPCServer.Dispose();
                    _OPCServer = null;
                }

                if (_OPCServer == null) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "_OPCServer is null, creating new object");
                    _OPCServer = new Opc.Hda.Server(fact, null);
                } 

                if (!_OPCServer.IsConnected) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "OPC server is disconnected, trying to connect");
                    _OPCServer.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential(), null));
                    if (!_OPCServer.IsConnected) {
                        _trace.TraceEvent(TraceEventType.Error, 0, "Connection failed without exception: {0}", url.ToString());
                        return false;
                    }
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Succesfully connected to {0}, obj: {1}", url.ToString(), _OPCServer.GetHashCode().ToString());
                }

                Status = _OPCServer.GetStatus();
                _trace.TraceEvent(TraceEventType.Verbose, 0, "OPC server status:\n"+
                    "\tCurrentTime:     {0}\n"+
                    "\tMaxReturnValues: {1}\n" +
                    "\tProductVersion:  {2}\n" +
                    "\tServerState:     {3}\n" +
                    "\tStartTime:       {4}\n" +
                    "\tStatusInfo:      {5}\n" +
                    "\tVendorInfo:      {6}\n", 
                    Status.CurrentTime,
                    Status.MaxReturnValues,
                    Status.ProductVersion,
                    Status.ServerState,
                    Status.StartTime,
                    Status.StatusInfo,
                    Status.VendorInfo);

                _trace.TraceEvent(TraceEventType.Verbose, 0, "SupportedAggregates:");
                SupportedAggregates = _OPCServer.GetAggregates();
                foreach (Opc.Hda.Aggregate agg in SupportedAggregates) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "{0}\t{1}\t{2}", agg.ID, agg.Name, agg.Description);
                }


            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Connection failed: {0}, {1}", url.ToString(), e.Message);
                return false;
            }

            return true;
        }


        // Checks (validates) given list of tags, and removes non-existent tags from the list
        // Error messages are written to console.
        // Returns true if all tags are valid, false otherwise.
        public bool Validate(List<string> Tagnames) {
            // Check if tags exist
            var ItemIds = new Opc.ItemIdentifier[Tagnames.Count()];
            for (int i = 0; i < Tagnames.Count(); i++) {
                ItemIds[i] = new Opc.ItemIdentifier(Tagnames[i]);
            }
            var res = _OPCServer.ValidateItems(ItemIds);
            bool all_valid = true;
            for (int i = Tagnames.Count() - 1; i >= 0; i--) {
                if (!res[i].ResultID.Succeeded()) {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Tag {0} is not valid and will not be read. Result_ID={1}", Tagnames[i], res[i].ResultID.ToString());
                    Console.ResetColor();
                    Tagnames.RemoveAt(i);
                    all_valid = false;
                }
            }
            return all_valid;
        }

        // This method uses Opc.Hda.Server class to read data
        public bool Read(string strStartTime,
                         string strEndTime,
                         string[] Tagnames,
                         int AggregateID,
                         int MaxValues,
                         int ResampleInterval,
                         bool IncludeBounds,
                         bool read_raw,
                         out Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {

            if ((Tagnames == null) || (Tagnames.Count() == 0)) {
                _trace.TraceEvent(TraceEventType.Error, 0, "HDAClient.Read: no tagnames to read");
                OPCHDAItemValues = null;
                return false;
            }

            if (!read_raw && !(SupportedAggregates.Any(a => a.ID == AggregateID))) {
                _trace.TraceEvent(TraceEventType.Error, 0, "HDAClient.Read: this aggregate is not supported");
                OPCHDAItemValues = null;
                return false;
            }

            var ItemIds = new Opc.ItemIdentifier[Tagnames.Count()];
            for (int i = 0; i < Tagnames.Count(); i++) {
                ItemIds[i] = new Opc.ItemIdentifier(Tagnames[i]);
            }
            // When using Server.ReadProcessed instead of Trend.ReadProcessed, we have to manually Server.CreateItems before.
            // We also have to call ReleaseItems after use
            var ItemIDResults = _OPCServer.CreateItems(ItemIds);
            for (int i = 0; i < Tagnames.Count(); i++) {
                if (!ItemIDResults[i].ResultID.Succeeded()) {
                    _trace.TraceEvent(TraceEventType.Error, 0, "Tag {0} from the read list is still not valid! Result_ID={1}", Tagnames[i], ItemIDResults[i].ResultID.ToString());
                }
            }

            DateTime dtStartTime, dtEndTime;
            Opc.Hda.Time hdaStartTime, hdaEndTime;
            ConvertStrDatetimeToHDADatetime(strStartTime, out dtStartTime, out hdaStartTime);
            ConvertStrDatetimeToHDADatetime(strEndTime, out dtEndTime, out hdaEndTime);

            // Specified dtStartTime may be > than dtEndTime, this means that data should be returned in reverse order (see spec)
            int order = 1;
            if (dtStartTime > dtEndTime)
                order = -1;

            // OPCTrend.MaxValues = MaxValues; // _OPCServer has no such property, it is passed as parameter to ReadRaw
            
            OPCHDAItemValues = null;

            try {
                if (read_raw) {
                    if (MaxValues > Status.MaxReturnValues) {
                        _trace.TraceEvent(TraceEventType.Warning, 0, "MaxValue was set to {0} (server cannot return more).", Status.MaxReturnValues);
                        MaxValues = Status.MaxReturnValues;
                    }
                    OPCHDAItemValues = _OPCServer.ReadRaw(hdaStartTime, hdaEndTime, MaxValues, IncludeBounds, ItemIDResults);
                } else {
                    var Items = new Opc.Hda.Item[Tagnames.Count()];
                    for (int i = 0; i < Tagnames.Count(); i++) {
                        Items[i] = new Opc.Hda.Item(ItemIDResults[i]);
                        Items[i].AggregateID = AggregateID;
                    }

                    Opc.Hda.ItemValueCollection[] tmpOPCHDAItemValues;

                    // Server returns data including start datetime, excluding last datetime: [dtStartPortion, dtEndPortion)
                    DateTime dtStartPortion = dtStartTime;
                    DateTime dtEndPortion;

                    while (order * ((dtEndTime - dtStartPortion).TotalSeconds) > 0) {
                        // is 32-bit int enough for Seconds?
                        int numvalues =  (int)  Math.Abs((dtEndTime - dtStartPortion).TotalSeconds / ResampleInterval);
                        if (numvalues > Status.MaxReturnValues) {
                            dtEndPortion = dtStartPortion.AddSeconds(ResampleInterval * Status.MaxReturnValues * order);
                            numvalues = Status.MaxReturnValues;
                        } else
                            dtEndPortion = dtEndTime;

                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Reading {0} values from {1} to {2}",
                            numvalues, dtStartPortion.ToString(), dtEndPortion.ToString());

                        tmpOPCHDAItemValues = _OPCServer.ReadProcessed(new Opc.Hda.Time(dtStartPortion),
                            new Opc.Hda.Time(dtEndPortion), ResampleInterval, Items);

                        if (dtStartPortion.Equals(dtStartTime)) {
                            OPCHDAItemValues = tmpOPCHDAItemValues;
                            //for (int i = 0; i < tmpOPCHDAItemValues.Count(); i++)
                            //    OPCHDAItemValues[i] = (Opc.Hda.ItemValueCollection)tmpOPCHDAItemValues[i].Clone();
                        } else {
                            for (int i = 0; i < tmpOPCHDAItemValues.Count(); i++)
                                OPCHDAItemValues[i].AddRange(tmpOPCHDAItemValues[i]);
                        }
                        dtStartPortion = dtEndPortion;
                    }
                }

                if (OPCHDAItemValues != null) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Number of tags = OPCHDAItemValues.Count()={0}", OPCHDAItemValues.Count());
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Number of values for the first tag = OPCHDAItemValues[0].Count={0}",
                        OPCHDAItemValues[0].Count);
                }
                return true;
            } catch (Opc.ResultIDException e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Opc.ResultIDException:" + e.ToString());
                // anyway, let's try to examine data
                if (OPCHDAItemValues == null) {
                    _trace.TraceEvent(TraceEventType.Error, 0, "OPCHDAItemValues == null");
                } else {
                    foreach (Opc.Hda.ItemValueCollection item in OPCHDAItemValues) {
                        _trace.TraceEvent(TraceEventType.Error, 0, "For tag {0}  the ResultID is {1}", item.ItemName, item.ResultID.ToString());
                    }
                }
                return false;
            } catch (Exception e) {
                _trace.TraceEvent(TraceEventType.Error, 0, e.Message);
                _trace.TraceEvent(TraceEventType.Error, 0, e.GetType().ToString());

                if (e.Data.Count > 0) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "  Extra details:");
                    foreach (System.Collections.DictionaryEntry de in e.Data)
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "    Key: {0,-20}      Value: {1}",
                                          "'" + de.Key.ToString() + "'", de.Value);
                }
                return false;
            } finally {
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Releasing items.");
                _OPCServer.ReleaseItems(ItemIds);
                for (int i = 0; i < Tagnames.Count(); i++) {
                    if (!ItemIDResults[i].ResultID.Succeeded()) {
                        _trace.TraceEvent(TraceEventType.Error, 0, "Tag {0} couldn't be released! Result_ID={1}", Tagnames[i], ItemIDResults[i].ResultID.ToString());
                    }
                }
            }
        }


        public void Disconnect() {
            if (_OPCServer != null) {
                if (_OPCServer.IsConnected) {
                    _OPCServer.Disconnect();
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Sucessfully disconnected from OPC HDA server.");
                } else {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "OPC HDA server is unexpectedly disconnected");
                }
                _OPCServer.Dispose();
            }
        }

        private void ConvertStrDatetimeToHDADatetime(string strDateTime, out DateTime dtDateTime, out Opc.Hda.Time hdaDateTime) {
            //Constructor Opc.Hda.Time(String) produces relative time, constructor Opc.Hda.Time(DateTime) produces absolute time. 
            //Constructor Opc.Hda.Time(String) doesn't parse the string. In case if time string is wrong, 
            //exception will be fired only when ReadProcessed is called.
            try {
                //Try to parse date and time. If it is in relative time format (for example NOW-30D),
                //exception will be generated
                dtDateTime = DateTime.Parse(strDateTime);
                //No exception => date and time in absolute format, pass them to Opc.Hda.Time constructor as DateTime
                hdaDateTime = new Opc.Hda.Time(dtDateTime);
            } catch (FormatException) {
                //Exception fired => Date and time in relative format.
                //Pass them to Opc.Hda.Time constructor as strings
                hdaDateTime = new Opc.Hda.Time(strDateTime);
                dtDateTime = hdaDateTime.ResolveTime();
            }

            _trace.TraceEvent(TraceEventType.Verbose, 0, "Timestamp {0} was recognized as {1}, IsRelative: {2}",
                strDateTime, hdaDateTime.ToString(), hdaDateTime.IsRelative);

            return;
        }
    }
}
