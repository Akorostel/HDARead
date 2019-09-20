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
        static TraceSource _trace = new TraceSource("HDAClientTraceSource");

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

        // Connect to OPC HDA server
        public bool Connect(string HostName, string ServerName) {

            var url = new Opc.URL();

            url.Scheme = Opc.UrlScheme.HDA;
            url.HostName = HostName;
            url.Path = ServerName;

            try {
                var fact = new OpcCom.Factory();
                if (_OPCServer == null) {
                    //TraceS("History2: _OPCServer Is Nothing, creating new object, trying to connect")
                    _OPCServer = new Opc.Hda.Server(fact, null);
                    _OPCServer.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential(), null));
                    if (_OPCServer.IsConnected) {
                        //TraceS("History2: succesfully connected to: " & url.ToString & ", obj: " & _OPCServer.GetHashCode().ToString)
                    }
                }
                //If connection was lost
                if (!_OPCServer.IsConnected) {
                    //TraceS("History2: OPC server is disconnected, trying to connect")
                    //Unfortunately, in case of lost connection simply calling .Connect() doesn't work :(
                    //Let's try to recreate the object from scratch
                    _OPCServer.Dispose();
                    _OPCServer = new Opc.Hda.Server(fact, null);
                    _OPCServer.Connect(url, new Opc.ConnectData(new System.Net.NetworkCredential(), null));
                    if (_OPCServer.IsConnected) {
                        //TraceS("History2: succesfully connected to: " & url.ToString & ", obj: " & _OPCServer.GetHashCode().ToString)
                    } else {
                        //TraceS("EXCEPTION", "History2: connection failed without exception: " & url.ToString, True)
                        return false;
                    }
                }
            } catch (Exception e) {
                //TraceS("EXCEPTION", "History2: connection failed: " & url.ToString & ", " & ex.Message, True)
                return false;
            }

            return true;
        }

        public bool Read(string StartTime,
                         string EndTime,
                         string[] Tagnames,
                         int AggregateID,
                         int MaxValues,
                         int ResampleInterval,
                         bool read_raw,
                         out Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            
            // Check if tags exist
            var ItemIds = new Opc.ItemIdentifier[Tagnames.Count()];
            for (int i = 0; i < Tagnames.Count(); i++) {
                ItemIds[i] =  new Opc.ItemIdentifier(Tagnames[i]);
            }
            var res = _OPCServer.ValidateItems(ItemIds);
            for (int i = 0; i < Tagnames.Count(); i++) {
                if (!res[i].ResultID.Succeeded()) {
                    _trace.TraceEvent(TraceEventType.Error, 0, "Tag {0} is not valid: Result_ID={1}, DiagnosticInfo={2}", Tagnames[i], res[i].ResultID.ToString(), res[i].DiagnosticInfo);
                }
            }

            var OPCTrend = new Opc.Hda.Trend(_OPCServer);
            //Constructor Opc.Hda.Time(String) produces relative time, constructor Opc.Hda.Time(DateTime) produces absolute time. 
            //Constructor Opc.Hda.Time(String) doesn't parse the string. In case if time string is wrong, 
            //exception will be fired only when ReadProcessed is called.

            try {
                DateTime StartDateTime;
                //Try to parse date and time. If it is in relative time format (for example NOW-30D),
                //exception will be generated
                StartDateTime = DateTime.Parse(StartTime);
                //No exception => date and time in absolute format, pass them to Opc.Hda.Time constructor as DateTime
                OPCTrend.StartTime = new Opc.Hda.Time(StartDateTime);
            } catch (FormatException e) {
                //Exception fired => Date and time in relative format.
                //Pass them to Opc.Hda.Time constructor as strings
                OPCTrend.StartTime = new Opc.Hda.Time(StartTime);
            }
            try {
                DateTime EndDateTime;
                EndDateTime = DateTime.Parse(EndTime);
                OPCTrend.EndTime = new Opc.Hda.Time(EndDateTime);
            } catch (FormatException e) {
                OPCTrend.EndTime = new Opc.Hda.Time(EndTime);
            }
            _trace.TraceEvent(TraceEventType.Verbose, 0, "From timestamp {0} was recognized as {1}, IsRelative: {2}", StartTime, OPCTrend.StartTime.ToString(), OPCTrend.StartTime.IsRelative);
            _trace.TraceEvent(TraceEventType.Verbose, 0, "To   timestamp {0} was recognized as {1}, IsRelative: {2}", EndTime, OPCTrend.EndTime.ToString(), OPCTrend.EndTime.IsRelative);

            OPCTrend.MaxValues = MaxValues;
            OPCTrend.ResampleInterval = ResampleInterval; // 0 - return just one value (see OPC HDA spec.)
            OPCTrend.Items.Clear();
            OPCHDAItemValues = null;
            
            try {
                for (int i = 0; i < Tagnames.Count(); i++) {
                    OPCTrend.AddItem(new Opc.ItemIdentifier(Tagnames[i]));
                    OPCTrend.Items[i].AggregateID = AggregateID;
                }
                if (read_raw)
                    OPCHDAItemValues = OPCTrend.ReadRaw();
                else
                    OPCHDAItemValues = OPCTrend.ReadProcessed();
                return true;
            } catch (Opc.ResultIDException e) {
                _trace.TraceEvent(TraceEventType.Error, 0, "Opc.ResultIDException:" + e.Message);
                _trace.TraceEvent(TraceEventType.Error, 0, "Opc.ResultIDException:" + e.ToString());
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
            }            
        }

        public void Disconnect() {
            if (_OPCServer != null) {
                if (_OPCServer.IsConnected)
                    _OPCServer.Disconnect();
                _OPCServer.Dispose();
            }
        }
    }
}
