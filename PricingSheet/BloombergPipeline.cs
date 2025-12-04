using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bloomberglp.Blpapi;

namespace PricingSheet
{
    public class BloombergPipeline
    {
        public List<string> Instruments { get; set; }
        public List<string> MaturityCodes { get; set; }
        public List<string> Fields { get; set; }

        public BloombergPipeline(List<string> Instruments, List<string> MaturityCodes, List<string> Fields ) 
        { 
            this.Instruments = Instruments; 
            this.MaturityCodes = MaturityCodes; 
            this.Fields = Fields;
        }

        public void Launch()
        {
            SessionOptions options = new SessionOptions
            {
                ServerHost = "localhost", // Bloomberg Terminal host
                ServerPort = 8194          // Default API port
            };

            using (Session session = new Session(options))
            {
                if (!session.Start())
                {
                    Console.WriteLine("Failed to start session");
                    return;
                }

                if (!session.OpenService("//blp/mktdata"))
                {
                    Console.WriteLine("Failed to open service");
                    return;
                }


                // Create subscriptions
                var subscriptions = new List<Subscription>();
                foreach (var instrument in Instruments)
                {
                    subscriptions.Add(new Subscription(instrument, Fields, new CorrelationID(instrument)));
                }

                // Subscribe
                session.Subscribe(subscriptions);
                Console.WriteLine("Subscribed to live data for multiple instruments/fields.");

                List<InstrumentData> instrumentDataList = new List<InstrumentData>();
                // Event loop
                while (true)
                {
                    Event ev = session.NextEvent();
                    foreach (Message msg in ev)
                    {
                        if (msg.MessageType.Equals("MarketDataEvents") || msg.MessageType.Equals("MarketDataUpdate"))
                        {
                            string instrument = msg.CorrelationID.ToString().Substring(6);
                            double bid = msg.HasElement("BID") ? msg.GetElementAsFloat64("BID") : double.NaN;
                            double ask = msg.HasElement("ASK") ? msg.GetElementAsFloat64("ASK") : double.NaN;

                            if (instrumentDataList.Where(x => x.Instrument == instrument).Any())
                                instrumentDataList.Where(x => x.Instrument == instrument).First().update(instrument, bid, ask);
                            else
                                instrumentDataList.Add(new InstrumentData(instrument, bid, ask));

                            //System.Threading.Thread.Sleep(1000);
                        }
                    }
                }
            }
        }

        private class InstrumentData
        {
            public string Instrument { get; set; }
            public double Bid { get; set; }
            public double Ask { get; set; }

            public InstrumentData()
            {
                Instrument = string.Empty;
                Bid = double.NaN;
                Ask = double.NaN;
            }

            public InstrumentData(string instrument, double bid, double ask)
            {
                Instrument = instrument;
                Bid = bid;
                Ask = ask;
            }

            public void update(string instrument, double bid, double ask)
            {
                if (!string.IsNullOrEmpty(instrument))
                    Instrument = instrument;
                if (!double.IsNaN(bid))
                    Bid = bid;
                if (!double.IsNaN(ask))
                    Ask = ask;
            }

        }

    }
}
