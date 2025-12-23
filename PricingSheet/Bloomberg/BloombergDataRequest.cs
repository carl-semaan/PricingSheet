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
        public List<string >Fields { get; set; }
        public MtM _MtM { get; set; }


        public BloombergDataRequest(MtM mtm, List<string> instruments, List<string> field)
        {
            _MtM = mtm;
            Instruments = instruments;
            Fields = field;
        }

        #region Fetch Prices 
        public async Task<List<UnderlyingSpotResponse>> FetchUlSpot()
        {
            Dictionary<string, UnderlyingSpotResponse> results = new Dictionary<string, UnderlyingSpotResponse>();

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

                    foreach(var field in Fields)
                        request.Append("fields", field);

                    session.SendRequest(request, null);

                    while (true)
                    {
                        Event ev = session.NextEvent(Constants.RequestTimeoutMS);

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

        private void ProcessMessage(Message msg, ref Dictionary<string, UnderlyingSpotResponse> results)
        {
            if (msg.MessageType == Name.GetName("ReferenceDataResponse"))
            {
                Element securityData = msg.GetElement("securityData");

                for (int i = 0; i < securityData.NumValues; i++)
                {
                    Element secData = securityData.GetValueAsElement(i);
                    string instrument = secData.GetElementAsString("security");

                    UnderlyingSpotResponse response = new UnderlyingSpotResponse
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

                        foreach (var field in Fields)
                        {
                            if (fieldData.HasElement(field))
                            {
                                response.Value =
                                    fieldData.GetElementAsFloat64(field);
                            }
                            else
                            {
                                response.Error = $"{Fields} not returned";
                            }
                        }
                    }

                    results[instrument] = response;
                }
            }
        }
        #endregion

        #region Fetch Instruments Info
        public async Task<List<InstrumentResponse>> FetchInstrument()
        {
            Dictionary<string, InstrumentResponse> results = new Dictionary<string, InstrumentResponse>();

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
                        request.Append("securities", instr);
                    }

                    foreach (var field in Fields)
                    {
                        request.Append("fields", field);
                    }

                    session.SendRequest(request, null);

                    while (true)
                    {
                        Event ev = session.NextEvent(Constants.RequestTimeoutMS);

                        if (ev.Type == Event.EventType.TIMEOUT)
                            throw new TimeoutException("Bloomberg request timed out");

                        foreach (Message msg in ev)
                            ProcessMessage(msg, ref results, Fields);

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

        private void ProcessMessage(Message msg, ref Dictionary<string, InstrumentResponse> results, List<string> fields)
        {
            if (msg.MessageType == Name.GetName("ReferenceDataResponse"))
            {
                Element securityData = msg.GetElement("securityData");

                for (int i = 0; i < securityData.NumValues; i++)
                {
                    Element secData = securityData.GetValueAsElement(i);
                    string instrument = secData.GetElementAsString("security");

                    InstrumentResponse response = new InstrumentResponse
                    {
                        Ticker = instrument
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

                        foreach(var field in fields)
                        {
                            if (fieldData.HasElement(field))
                            {
                                string value = fieldData.GetElementAsString(field);
                                switch (field)
                                {
                                    case "UNDERLYING":
                                        response.Underlying = value;
                                        break;
                                    case "SHORT_NAME":
                                        response.ShortName = value;
                                        break;
                                    case "EXCHANGE_CODE":
                                        response.ExchangeCode = value;
                                        break;
                                    case "INSTRUMENT_TYPE":
                                        response.InstrumentType = value;
                                        break;
                                    case "CRNCY":
                                        response.Currency = value;
                                        break;
                                    case "ICB_SUPER_SECTOR_NAME":
                                        response.ICBSuperSectorName = value;
                                        break;
                                }
                            }
                            else
                            {
                                response.Error += $"{field} not returned; ";
                            }
                        }
                    }

                    results[instrument] = response;
                }
            }
        }
        #endregion
    }

    public class UnderlyingSpotResponse : UnderlyingSpot
    {
        public string Error { get; set; }

        public UnderlyingSpotResponse() { }
        public UnderlyingSpotResponse(string underlying, double? value, string error) : base(underlying, value)
        {
            Error = error;
        }
    }

    public class InstrumentResponse : Instruments
    {
        public string Error { get; set; }
        public InstrumentResponse() { }
        public InstrumentResponse(string ticker, string underlying, string shortName, string exchangeCode, string instrumentType, string currency, string ICBSupersectorName, string error)
            : base(ticker, underlying, shortName, exchangeCode, instrumentType, currency, ICBSupersectorName)
        {
            Error = error;
        }
    }
}
