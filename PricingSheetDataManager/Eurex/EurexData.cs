using CsvHelper;
using DocumentFormat.OpenXml.Spreadsheet;
using PricingSheetCore;
using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PricingSheetCore.Readers;

namespace PricingSheetDataManager.Eurex
{
    public class EurexData
    {
        public static async Task<List<EurexInstruments>> FetchEurexInstruments()
        {
            Console.WriteLine("Fetching Eurex data...");
            // Fetch the eurex file
            string tempPath = await GetEurexFile();

            Console.WriteLine("Parsing Eurex data...");

            List<EurexInstruments> eurexInstruments = new List<EurexInstruments>();
            try
            {
                // Read the eurex file
                CSVReader csvReader = new CSVReader(Path.GetDirectoryName(tempPath), Path.GetFileName(tempPath), ";");
                eurexInstruments = csvReader.LoadClass<EurexInstruments>().Where(x => x.ProductGroup == "SINGLE STOCK DIVIDEND FUTURES").ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error occured while parsing Eurex Instruments: {e.Message}");
            }
            finally
            {
                // Clean up temp file
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }

            Console.WriteLine("Eurex data fetch and parse completed.");
            return eurexInstruments;
        }

        private static async Task<string> GetEurexFile()
        {
            IPage page = null;
            IBrowser browser = null;
            try
            {
                Console.WriteLine("Launching browser...");
                // Launching browser
                (browser, page) = await LaunchBrowser();

                Console.WriteLine("Navigating to Eurex URL...");
                // Navigating to Eurex URL
                await page.GoToAsync(Constants.EurexURL);

                Console.WriteLine("Downloading Eurex file...");
                // Downloading Eurex file
                string tempPath = await DownloadEurexFile(page);

                return tempPath;
            }
            catch (Exception e)
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

            throw new Exception("Eurex file retrieval failed");
        }

        private static async Task<(IBrowser, IPage)> LaunchBrowser()
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
                Timeout = Constants.ScraperTimeoutMS,
            };

            // Launching browser with retry mechanism-
            for (int i = 0; i < Constants.MaxAttempts; i++)
            {
                try
                {
                    IBrowser browser = await Puppeteer.LaunchAsync(launchOptions);
                    IPage page = await browser.NewPageAsync();

                    page.DefaultNavigationTimeout = Constants.ScraperTimeoutMS;
                    browser.DefaultWaitForTimeout = Constants.ScraperTimeoutMS;

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

        private static async Task<string> DownloadEurexFile(IPage page)
        {
            string temporaryPath = System.IO.Path.GetTempPath();

            await page.Client.SendAsync("Page.setDownloadBehavior", new
            {
                behavior = "allow",
                downloadPath = temporaryPath
            });

            for (int i = 0; i < Constants.MaxAttempts; i++)
            {
                try
                {
                    // Waiting for our selector to load
                    await page.WaitForSelectorAsync(".space-cta a");

                    // Loading all the buttons
                    var buttons = await page.QuerySelectorAllAsync(".space-cta a");

                    if (buttons == null || buttons.Length == 0)
                        throw new Exception("No download button found on Eurex page.");

                    // Clicking on the last button to download the file
                    await buttons[^1].ClickAsync();

                    // Checking if the file is downloading
                    string targetFile = await IsDownloadComplete(temporaryPath);

                    return targetFile;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {i + 1}/{Constants.MaxAttempts} to download Eurex file failed: {ex.Message}");
                    if (i == Constants.MaxAttempts - 1)
                        throw new Exception($"Failed to download Eurex file after {Constants.MaxAttempts} attempts. Error: {ex.Message}");
                    await Task.Delay(10000);
                }
            }

            throw new Exception("Eurex file download failed");
        }

        private static async Task<string> IsDownloadComplete(string temporaryPath)
        {
            for (int i = 0; i < Constants.MaxAttempts; i++)
            {
                var files = System.IO.Directory.GetFiles(temporaryPath, "*.csv");
                if (files.Length > 0)
                    return files.First();
                await Task.Delay(5000);
            }

            throw new Exception("File download did not complete in expected time.");
        }

    }
}
