using DocumentFormat.OpenXml.Spreadsheet;
using PricingSheetCore;
using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheetDataManager.Eurex
{
    public class Eurex
    {
        public async Task<List<EurexInstruments>> FetchEurexInstruments()
        {
            string tempPath = await GetEurexFile();

            return new List<EurexInstruments>();
        }

        private async Task<string> GetEurexFile()
        {
            IPage page = null;
            IBrowser browser = null;
            try
            {
                (browser, page) = await LaunchBrowser();

                await page.GoToAsync(Constants.EurexURL);
            }
            catch(Exception e)
            {
                Console.WriteLine($"Error occured while fetching Eurex Instruments: {e.Message}");
            }
            finally
            {
                if (browser != null)
                {
                    await browser.CloseAsync();
                }
            }

            return string.Empty;
        }

        private async Task<(IBrowser, IPage)> LaunchBrowser()
        {
            // Setting browser options
            var launchOptions = new LaunchOptions
            {
                Headless = true,
                IgnoredDefaultArgs = new[] { "--enable-automation" },
                ExecutablePath = @"G:\Shared drives\Arbitrage\Tools\31.MonthlyPresentation\Chrome\Application\chrome.exe",
                Args = new[]
                    {
                        "--start-maximized"
                    },
                DefaultViewport = null,
                Timeout = 120000,
            };

            // Launching browser with retry mechanism-
            for (int i = 0; i < Constants.MaxAttempts; i++)
            {
                try
                {
                    IBrowser browser = await Puppeteer.LaunchAsync(launchOptions);
                    IPage page = await browser.NewPageAsync();

                    return (browser, page); 
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {i + 1}/{Constants.MaxAttempts} to launch browser failed: {ex.Message}");
                    if (i == Constants.MaxAttempts - 1)
                        throw new Exception($"Failed to launch browser after {Constants.MaxAttempts} attempts. Error: {ex.Message}");
                    await Task.Delay(10000);
                }
            }
            throw new Exception("Broswer launch failed");
        }
    }
}
