using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bloomberglp.Blpapi;
using ExcelVSTO = Microsoft.Office.Tools.Excel;

namespace PricingSheet
{
    public class BloombergPipeline
    {
        private CancellationToken _token;
        private SynchronizationContext _syncContext = SynchronizationContext.Current;
        public List<Flux.Instruments> Instruments { get; set; }
        public List<string> MaturityCodes { get; set; }
        public List<string> Fields { get; set; }
        public ExcelVSTO.Worksheet Sheet { get; set; }

        public BloombergPipeline(ExcelVSTO.Worksheet Sheet, List<Flux.Instruments> Instruments, List<string> MaturityCodes, List<string> Fields)
        {
            this.Sheet = Sheet;
            this.Instruments = Instruments;
            this.MaturityCodes = MaturityCodes;
            this.Fields = Fields;
        }

        public void Launch(CancellationToken token)
        {
            _token = token;

            SessionOptions options = new SessionOptions
            {
                ServerHost = "localhost", // Bloomberg Terminal host
                ServerPort = 8194          // Default API port
            };

            try
            {
                using (Session session = new Session(options))
                {
                    if (!session.Start())
                        throw new Exception("Failed to start session");

                    if (!session.OpenService("//blp/mktdata"))
                        throw new Exception("Failed to open service");

                    // Create subscriptions
                    var subscriptions = GetSubscriptions();

                    // Subscribe
                    session.Subscribe(subscriptions);
                    System.Diagnostics.Debug.WriteLine("Subscribed to live data for multiple instruments/fields.");

                    // Event loop
                    int ctr = 0;
                    while (!_token.IsCancellationRequested)
                    {
                        try {
                            Log($"Waiting for next event... {ctr} {token.IsCancellationRequested}");
                            Event ev = session.NextEvent();
                            Log($"Received event: {ev.Type}");
                            foreach (Message msg in ev)
                            {
                                Log($"Event Type: {ev.Type},\nMessage: {msg}");
                                if (msg.MessageType.Equals("MarketDataEvents") || msg.MessageType.Equals("MarketDataUpdate"))
                                {
                                    string instrument = msg.CorrelationID.ToString().Substring(6);
                                    foreach (var field in Fields)
                                    {
                                        if (msg.HasElement(field))
                                        {
                                            var element = msg.GetElement(field);
                                            if (element.Datatype == Bloomberglp.Blpapi.Schema.Datatype.FLOAT64 || element.Datatype == Bloomberglp.Blpapi.Schema.Datatype.INT32)
                                            {
                                                var value = msg.GetElementAsFloat64(field);
                                                if (!double.IsNaN(value))
                                                    updateSheet(instrument, field, value);
                                            }
                                        }
                                    }
                                }
                                else if(msg.MessageType.Equals("SubscriptionFailure") || msg.MessageType.Equals("SubscriptionTerminated"))
                                {
                                    Log($"Subscription failure for {msg.CorrelationID}: {msg}");
                                    return;
                                }
                            }
                            ctr++;
                        }
                        catch (Exception ex)
                        {
                            Log($"Loop exception: {ex}");
                            System.Diagnostics.Debug.WriteLine($"Error: {ex.ToString()}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"outer exception: {ex}");
                System.Diagnostics.Debug.WriteLine($"Error: {ex.ToString()}");
            }
        }

        private void Log(string message)
        {
            File.AppendAllText(@"C:\Users\\Melanion\Desktop\Log.txt", $"{DateTime.Now}: {message}{Environment.NewLine}");
        }

        private List<Subscription> GetSubscriptions()
        {
            List<Subscription> subscriptions = new List<Subscription>();
            string instr;

            foreach (var instrument in Instruments)
            {
                foreach (var maturity in MaturityCodes)
                {
                    instr = $"{instrument.Ticker}={maturity} {instrument.ExchangeCode} {instrument.InstrumentType}";
                    subscriptions.Add(new Subscription(instr, Fields, new CorrelationID(instr)));
                }
            }

            return subscriptions;
        }

        private void updateSheet(string instrument, string field, double value)
        {
            string[] parts = instrument.Split('=');

            (int row, int col) = Utils.FindCellFlux(parts[1].Split(' ')[0], field, parts[0]);

            SheetDisplay.DisplayCell(Sheet, new DataCell(value.ToString(), IsCentered: true), row, col);
        }

        //double bid = msg.HasElement("BID") ? msg.GetElementAsFloat64("BID") : double.NaN;
        //double ask = msg.HasElement("ASK") ? msg.GetElementAsFloat64("ASK") : double.NaN;

        //if (!double.IsNaN(bid) || !double.IsNaN(ask))
        //    System.Diagnostics.Debug.WriteLine($"Instrument: {instrument}, Bid: {bid}, Ask: {ask}");

        //if (instrumentDataList.Where(x => x.Instrument == instrument).Any())
        //    instrumentDataList.Where(x => x.Instrument == instrument).First().update(instrument, bid, ask);
        //else
        //    instrumentDataList.Add(new InstrumentData(instrument, bid, ask));

        //System.Threading.Thread.Sleep(1000);


        //private class InstrumentData
        //{
        //    public string Instrument { get; set; }
        //    public double Bid { get; set; }
        //    public double Ask { get; set; }

        //    public InstrumentData()
        //    {
        //        Instrument = string.Empty;
        //        Bid = double.NaN;
        //        Ask = double.NaN;
        //    }

        //    public InstrumentData(string instrument, double bid, double ask)
        //    {
        //        Instrument = instrument;
        //        Bid = bid;
        //        Ask = ask;
        //    }

        //    public void update(string instrument, double bid, double ask)
        //    {
        //        if (!string.IsNullOrEmpty(instrument))
        //            Instrument = instrument;
        //        if (!double.IsNaN(bid))
        //            Bid = bid;
        //        if (!double.IsNaN(ask))
        //            Ask = ask;
        //    }

        //}

    }
}
