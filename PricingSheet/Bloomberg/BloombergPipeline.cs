using Bloomberglp.Blpapi;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static PricingSheet.Flux;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using ExcelVSTO = Microsoft.Office.Tools.Excel;
using PricingSheetCore.Models;

namespace PricingSheet
{
    public class BloombergPipeline
    {
        private CancellationToken _token;

        public List<Instruments> Instruments { get; set; }
        public List<string> MaturityCodes { get; set; }
        public List<string> Fields { get; set; }
        public ExcelVSTO.Worksheet Sheet { get; set; }
        private Flux FluxInstance = Flux.FluxInstance;
        private ConcurrentQueue<Event> _eventQueue = new ConcurrentQueue<Event>();
        private MtM MtMInstance = MtM.MtMInstance;
        private Univ UnivInstance = Univ.UnivInstance;
        private Ribbons.Ribbon RibbonInstance;
        public BloombergPipeline(ExcelVSTO.Worksheet Sheet, List<Instruments> Instruments, List<string> MaturityCodes, List<string> Fields)
        {
            this.Sheet = Sheet;
            this.Instruments = Instruments;
            this.MaturityCodes = MaturityCodes;
            this.Fields = Fields;
        }

        public void Launch(CancellationToken token)
        {
            // Waiting for MtM files to load
            MtM.MtMInstance.FilesLoaded.Wait();

            // Setting the Ribbon Instance
            RibbonInstance = Ribbons.Ribbon.RibbonInstance;

            // Setting Ribbon Data
            RibbonInstance?.SetStatus(bbgStatus: "Connecting...");
            RibbonInstance?.SetActiveSubscription(0);

            // Initializing the cancellation token
            _token = token;

            // Setting Session Options
            SessionOptions options = new SessionOptions
            {
                ServerHost = "localhost",
                ServerPort = 8194
            };

            try
            {
                using (Session session = new Session(options))
                {
                    // Starting Session
                    if (!session.Start())
                        throw new Exception("Failed to start session");

                    // Opening Market Data Service
                    if (!session.OpenService("//blp/mktdata"))
                        throw new Exception("Failed to open service");

                    // Setting the Ribbon Status to Live
                    RibbonInstance?.SetStatus(bbgStatus: "Live");

                    //  Subscribing to live data
                    var subscriptions = GetSubscriptions();

                    if (subscriptions.Count > Constants.MaxActiveInstruments)
                    {
                        throw new Exception($"Number of subscriptions ({subscriptions.Count}) exceeds the maximum allowed ({Constants.MaxActiveInstruments}).");
                    }
                    else if (subscriptions.Count == 0)
                    {
                        RibbonInstance?.SetStatus(bbgStatus: "Pending");
                        RibbonInstance?.SetActiveSubscription(subscriptions.Count);
                        return;
                    }

                    session.Subscribe(subscriptions);
                    System.Diagnostics.Debug.WriteLine("Subscribed to live data.");

                    // Updating Ribbon with active subscriptions count
                    RibbonInstance?.SetActiveSubscription(subscriptions.Count);

                    // Setting up consumer threads
                    var consumerThreads = new List<Thread>();
                    int consumerthreadsCount = Environment.ProcessorCount;
                    for (int i = 0; i < consumerthreadsCount; i++)
                    {
                        var consumerThread = new Thread(() => ProcessEvents(token));
                        consumerThread.Start();
                        consumerThreads.Add(consumerThread);
                    }

                    // Reading events from Bloomberg and enqueueing them
                    while (!token.IsCancellationRequested)
                    {
                        Event ev = session.NextEvent(Constants.TimeoutMS);

                        if (ev.Type == Event.EventType.TIMEOUT)
                            continue;

                        _eventQueue.Enqueue(ev);
                    }

                    // Setting Ribbon Status to Disconnected
                    RibbonInstance?.SetStatus(bbgStatus: "Disconnected");

                    // Waiting for consumer threads to finish
                    foreach (var consumerThread in consumerThreads)
                        consumerThread.Join();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error: {ex}");
                RibbonInstance?.SetStatus(bbgStatus: "Failed");
            }
        }

        private void ProcessEvents(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                if (_eventQueue.TryDequeue(out Event ev))
                {
                    foreach (Message msg in ev)
                    {
                        try
                        {
                            HandleMessage(msg);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Processing error: {ex}");
                        }
                    }
                }
                else
                {
                    Thread.Sleep(1);
                }
            }
        }

        private void HandleMessage(Message msg)
        {
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
                            {
                                FluxInstance.UpdateMatrixSafe(instrument, field, value);
                                UnivInstance.UpdateMatrixSafe(instrument, field, value);
                            }
                        }
                    }
                }
            }
            else if (msg.MessageType.Equals("SubscriptionFailure") || msg.MessageType.Equals("SubscriptionTerminated") || msg.MessageType.Equals("SessionTerminated"))
            {
                System.Diagnostics.Debug.WriteLine($"Subscription failure for {msg.CorrelationID}: {msg}");
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
            // Waiting for MtM files to load 
            MtMInstance.FilesLoaded.Wait();

            // Setting the Ribbon Instance
            RibbonInstance = Ribbons.Ribbon.RibbonInstance;

            RibbonInstance?.SetStatus(bbgStatus: "Offline");

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
                            FluxInstance.UpdateMatrixSafe(instr, field, value);
                            UnivInstance.UpdateMatrixSafe(instr, field, value);
                        }
                    }
                }
                Thread.Sleep(Constants.ThreadSleep);
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
