using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bloomberglp.Blpapi;
using ExcelVSTO = Microsoft.Office.Tools.Excel;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace PricingSheet
{
    public class BloombergPipeline
    {
        private static readonly int timeoutMs = 5000;
        private CancellationToken _token;

        public List<Flux.Instruments> Instruments { get; set; }
        public List<string> MaturityCodes { get; set; }
        public List<string> Fields { get; set; }
        public ExcelVSTO.Worksheet Sheet { get; set; }
        private Flux _Flux;

        public BloombergPipeline(Flux Flux, ExcelVSTO.Worksheet Sheet, List<Flux.Instruments> Instruments, List<string> MaturityCodes, List<string> Fields)
        {
            _Flux = Flux;
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
                ServerHost = "localhost",
                ServerPort = 8194        
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
                    while (!_token.IsCancellationRequested)
                    {
                        try
                        {
                            System.Diagnostics.Debug.WriteLine($"Waiting for next event... ");
                            Event ev = session.NextEvent(timeoutMs);
                            if (ev.Type == Event.EventType.TIMEOUT)
                            {
                                System.Diagnostics.Debug.WriteLine("Timeout event received, continuing...");
                                continue;
                            }

                            System.Diagnostics.Debug.WriteLine($"Received event: {ev.Type}");
                            foreach (Message msg in ev)
                            {
                                System.Diagnostics.Debug.WriteLine($"Event Type: {ev.Type},\nMessage: {msg}");
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
                                                    _Flux.UpdateMatrixSafe(instrument, field, value);
                                            }
                                        }
                                    }
                                }
                                else if (msg.MessageType.Equals("SubscriptionFailure") || msg.MessageType.Equals("SubscriptionTerminated") || msg.MessageType.Equals("SessionTerminated"))
                                {
                                    System.Diagnostics.Debug.WriteLine($"Subscription failure for {msg.CorrelationID}: {msg}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error: {ex.ToString()}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error: {ex.ToString()}");
            }
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

        public void LaunchOfflineTest(CancellationToken token)
        {
            _token = token;
            Random rand = new Random();
            while (!_token.IsCancellationRequested)
            {
                foreach (var instrument in Instruments)
                {
                    foreach (var maturity in MaturityCodes)
                    {
                        string instr = $"{instrument.Ticker}={maturity} {instrument.ExchangeCode} {instrument.InstrumentType}";
                        foreach (var field in Fields)
                        {
                            double value = Math.Round(rand.NextDouble() * 100, 2);
                            _Flux.UpdateMatrixSafe(instr, field, value);
                        }
                    }
                }
                Thread.Sleep(1000);
            }
        } 

        private void updateSheet(string instrument, string field, double value)
        {
            string[] parts = instrument.Split('=');

            (int row, int col) = Utils.FindCellFlux(parts[1].Split(' ')[0], field, parts[0]);

            SheetDisplay.DisplayCell(Sheet, new DataCell(value.ToString(), IsCentered: true), row, col);
        }
    }
}
