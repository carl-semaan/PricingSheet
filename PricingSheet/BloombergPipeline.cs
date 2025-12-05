using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bloomberglp.Blpapi;

namespace PricingSheet
{
    public class BloombergPipeline
    {
        private CancellationToken _token;
        public List<Flux.Instruments> Instruments { get; set; }
        public List<string> MaturityCodes { get; set; }
        public List<string> Fields { get; set; }

        public BloombergPipeline(List<Flux.Instruments> Instruments, List<string> MaturityCodes, List<string> Fields)
        {
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
                    Console.WriteLine("Subscribed to live data for multiple instruments/fields.");

                    // Event loop
                    while (!_token.IsCancellationRequested)
                    {
                        Event ev = session.NextEvent();
                        foreach (Message msg in ev)
                        {
                            if (msg.MessageType.Equals("MarketDataEvents") || msg.MessageType.Equals("MarketDataUpdate"))
                            {
                                string instrument = msg.CorrelationID.ToString().Substring(6);
                                double bid = msg.HasElement("BID") ? msg.GetElementAsFloat64("BID") : double.NaN;
                                double ask = msg.HasElement("ASK") ? msg.GetElementAsFloat64("ASK") : double.NaN;

                                if (!double.IsNaN(bid) || !double.IsNaN(ask))
                                    Console.WriteLine($"Instrument: {instrument}, Bid: {bid}, Ask: {ask}");

                                //if (instrumentDataList.Where(x => x.Instrument == instrument).Any())
                                //    instrumentDataList.Where(x => x.Instrument == instrument).First().update(instrument, bid, ask);
                                //else
                                //    instrumentDataList.Add(new InstrumentData(instrument, bid, ask));

                                //System.Threading.Thread.Sleep(1000);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
