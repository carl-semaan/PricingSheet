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
using System.Collections.Concurrent;

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
        private ConcurrentQueue<Event> _eventQueue = new ConcurrentQueue<Event>();
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
            MtM.MtMInstance.FilesLoaded.Wait();

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

                    var subscriptions = GetSubscriptions();
                    session.Subscribe(subscriptions);
                    System.Diagnostics.Debug.WriteLine("Subscribed to live data.");

                    var consumerThreads = new List<Thread>();
                    int consumerthreadsCount = Environment.ProcessorCount;
                    for (int i = 0; i < consumerthreadsCount; i++)
                    {
                        var consumerThread = new Thread(() => ProcessEvents(token));
                        consumerThread.Start();
                        consumerThreads.Add(consumerThread);
                    }

                    while (!token.IsCancellationRequested)
                    {
                        Event ev = session.NextEvent(timeoutMs);
                        _eventQueue.Enqueue(ev);
                    }

                    foreach (var consumerThread in consumerThreads)
                        consumerThread.Join();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error: {ex}");
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
