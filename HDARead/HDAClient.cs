using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Opc;
using OpcCom;

namespace HDARead {
    class HDAClient {

        private Opc.Hda.Server _OPCServer = null;

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

        public bool Read(string StartTime, string EndTime, List<string> Tagnames, int AggregateID) {
            var OPCTrend = new Opc.Hda.Trend(_OPCServer);
            int i;
            //Constructor Opc.Hda.Time(String) produces relative time, constructor Opc.Hda.Time(DateTime) produces absolute time. 
            //Constructor Opc.Hda.Time(String) doesn't parse the string. In case if time string is wrong, 
            //exception will be fired only when ReadProcessed is called.
            try {
                DateTime StartDateTime, EndDateTime;
                //Try to parse date and time. If it is in relative time format (for example NOW-30D),
                //exception will be generated
                StartDateTime = DateTime.Parse(StartTime);
                EndDateTime = DateTime.Parse(EndTime);
                //No exception => date and time in absolute format, pass them to Opc.Hda.Time constructor as DateTime
                OPCTrend.StartTime = new Opc.Hda.Time(StartDateTime);
                OPCTrend.EndTime = new Opc.Hda.Time(EndDateTime);
            } catch (Exception e) {
                //Exception fired => Date and time in relative format.
                //Pass them to Opc.Hda.Time constructor as strings
                OPCTrend.StartTime = new Opc.Hda.Time(StartTime);
                OPCTrend.EndTime = new Opc.Hda.Time(EndTime);
            }

            //Debug.Print("From " & StartTime.Val & " " & OPCTrend.StartTime.ToString & ", IsRelative: " & OPCTrend.StartTime.IsRelative)
            //Debug.Print("To " & EndTime.Val & " " & OPCTrend.EndTime.ToString & ", IsRelative: " & OPCTrend.EndTime.IsRelative)

            OPCTrend.MaxValues = 10;
            OPCTrend.ResampleInterval = 0; // return just one value (see OPC HDA spec.)

            Opc.Hda.ItemValueCollection OPCHDAItemValues[] = null;

            foreach (string tag in Tagnames) {
                OPCTrend.Items.Clear();
                try {
                    OPCTrend.AddItem(new Opc.ItemIdentifier(tag));
                    OPCTrend.Items[0].AggregateID = AggregateID;

                    OPCHDAItemValues = OPCTrend.ReadProcessed();

                    //Debug.Print(OPCHDAItemValues(0).ItemName, OPCHDAItemValues(0).ResultID.Name)
                    if ((OPCHDAItemValues[0][0].Value!=null) && (OPCHDAItemValues[0].ResultID == Opc.ResultID.S_OK) {
                        ArrResults(i) = CSng(OPCHDAItemValues(0)(0).Value)
                        'Convert timestamp according to OutputTimestampFormat
                        If OutputTimestampFormat.Val = "" Then
                            ArrTimestamps(i) = OPCHDAItemValues(0)(0).Timestamp.ToString
                        Else
                            ArrTimestamps(i) = OPCHDAItemValues(0)(0).Timestamp.ToString(OutputTimestampFormat.Val)
                        End If
                        ArrQualities(i) = OPCHDAItemValues(0)(0).Quality.ToString
                    } else {
                        //TraceS("History2: item " & i & " " & ArrTagnames(i) & ": ResultID=" & OPCHDAItemValues(0).ResultID.ToString)
                        ArrResults(i) = Single.NaN
                        ArrTimestamps(i) = ""
                        ArrQualities(i) = "ERROR"
                    }
                 } catch (Exception e) {
                    //TraceS("EXCEPTION", "History2: item " & i & " " & ArrTagnames(i) & ": " & ex.Message, True)
                    /*'TraceException("History2: item " & i & " " & ArrTagnames(i), ex, True)
                    If ArrUseLGV(i) = urtNOYES.uNO Then
                        ArrResults(i) = Single.NaN
                        ArrTimestamps(i) = ""
                        ArrQualities(i) = "ERROR"
                    End If*/
                 }
            }

            return true;
        }

        /*




            'Write arrays
            Results.PutArray(ArrResults, sErr)
            Timestamps.PutArray(ArrTimestamps, sErr)
            Qualities.PutArray(ArrQualities, sErr)
        */
    }
}
