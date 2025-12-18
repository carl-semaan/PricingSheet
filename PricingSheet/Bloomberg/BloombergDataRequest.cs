using Bloomberglp.Blpapi;
using PricingSheet.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Bloomberg
{
    public class BloombergDataRequest
    {
        public List<string> Instruments { get; set; }
        public string Field { get; set; }
        public MtM _MtM { get; set; }


        public BloombergDataRequest(MtM mtm, List<string> instruments, string field)
        {
            _MtM = mtm;
            Instruments = instruments;
            Field = field;
        }

        public async Task<List<APIresponse>> FetchData()
        {
            Dictionary<string, APIresponse> results = new Dictionary<string, APIresponse>();

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

                    if (!session.OpenService("//blp/refdata"))
                        throw new Exception("Failed to open service");

                    Service reDataService = session.GetService("//blp/refdata");

                    Request request = reDataService.CreateRequest("ReferenceDataRequest");

                    foreach (var instr in Instruments)
                    {
                        try
                        {
                            request.Append("securities", instr);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error adding instrument {instr}: {ex}");
                        }
                    }

                    request.Append("fields", Field);

                    session.SendRequest(request, null);

                    while (true)
                    {
                        Event ev = session.NextEvent(Constants.TimeoutMS);

                        if (ev.Type == Event.EventType.TIMEOUT)
                            throw new TimeoutException("Bloomberg request timed out");

                        foreach (Message msg in ev)
                            ProcessMessage(msg, ref results);

                        if (ev.Type == Event.EventType.RESPONSE)
                            break;
                    }

                    session.Stop();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error: {ex}");
            }

            return results.Values.ToList();
        }

        private void ProcessMessage(Message msg, ref Dictionary<string, APIresponse> results)
        {
            if (msg.MessageType == Name.GetName("ReferenceDataResponse"))
            {
                Element securityData = msg.GetElement("securityData");

                for (int i = 0; i < securityData.NumValues; i++)
                {
                    Element secData = securityData.GetValueAsElement(i);
                    string instrument = secData.GetElementAsString("security");

                    APIresponse response = new APIresponse
                    {
                        Underlying = instrument
                    };

                    if (secData.HasElement("securityError"))
                    {
                        response.Error = secData
                            .GetElement("securityError")
                            .GetElementAsString("message");
                    }
                    else
                    {
                        Element fieldData = secData.GetElement("fieldData");

                        if (fieldData.HasElement(Field))
                        {
                            response.Value =
                                fieldData.GetElementAsFloat64(Field);
                        }
                        else
                        {
                            response.Error = $"{Field} not returned";
                        }
                    }

                    results[instrument] = response;
                }
            }
        }
    }

    public class APIresponse : UnderlyingSpot
    {
        public string Error { get; set; }

        public APIresponse() { }
        public APIresponse(string underlying, double? value, string error) : base(underlying, value)
        {
            Error = error;
        }
    }
}
