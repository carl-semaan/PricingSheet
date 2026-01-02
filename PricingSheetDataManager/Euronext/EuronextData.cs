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

namespace PricingSheetDataManager.Euronext
{
    public class EuronextData
    {
        public static async Task<List<EuronextInstruments>> FetchEuronextInstruments()
        {
            Console.WriteLine("Fetching Euronext data...");
            // Fetch the Euronext file
            string tempPath = await GetEuronextFile();

            Console.WriteLine("Parsing Euronext data...");

            List<EuronextInstruments> EuronextInstruments = new List<EuronextInstruments>();
            try
            {
                // Wait for file to be fully written
                await Task.Delay(2000);
                // Read the Euronext file
                CSVReader csvReader = new CSVReader(Path.GetDirectoryName(tempPath), Path.GetFileName(tempPath), Delimiter: ";", SkipFirstRow: true);
                EuronextInstruments = csvReader.LoadClass<EuronextInstruments>().Where(x => x.ProductFamily == "Dividend Stock Futures").ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error occured while parsing Euronext Instruments: {e.Message}");
            }
            finally
            {
                // Clean up temp file
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }

            Console.WriteLine("Euronext data fetch and parse completed.");
            return EuronextInstruments;
        }

        private static async Task<string> GetEuronextFile()
        {
            IPage page = null;
            IBrowser browser = null;
            try
            {
                Console.WriteLine("Launching browser...");
                // Launching browser
                (browser, page) = await LaunchBrowser();

                Console.WriteLine("Navigating to Euronext URL...");
                // Navigating to Euronext URL
                await page.GoToAsync(Constants.EuronextURL);

                Console.WriteLine("Downloading Euronext file...");
                // Downloading Euronext file
                string tempPath = await DownloadEuronextFile(page);

                return tempPath;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error occured while fetching Euronext Instruments: {e.Message}");
            }
            finally
            {
                if (browser != null)
                {
                    await browser.CloseAsync();
                }
            }

            throw new Exception("Euronext file retrieval failed");
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

        private static async Task<string> DownloadEuronextFile(IPage page)
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
                    await page.WaitForSelectorAsync(".view-header button");

                    // Loading all the buttons
                    var buttons = await page.QuerySelectorAllAsync(".view-header button");

                    if (buttons == null || buttons.Length == 0)
                        throw new Exception("No download button found on Euronext page.");

                    // Clicking on the download button
                    await buttons[^1].ClickAsync();

                    // Waiting for the pop over to appear
                    await page.WaitForSelectorAsync(".popover-body div");

                    // Loading all the buttons
                    var targetDivs = await page.QuerySelectorAllAsync(".popover-body div");

                    if (targetDivs == null || targetDivs.Length == 0)
                        throw new Exception("No download option found on Euronext page.");

                    // Clicking on the last div to download the file
                    await targetDivs[0].ClickAsync();

                    // Checking if the file is downloading
                    string targetFile = await IsDownloadComplete(temporaryPath);

                    return targetFile;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {i + 1}/{Constants.MaxAttempts} to download Euronext file failed: {ex.Message}");
                    if (i == Constants.MaxAttempts - 1)
                        throw new Exception($"Failed to download Euronext file after {Constants.MaxAttempts} attempts. Error: {ex.Message}");
                    await Task.Delay(10000);
                }
            }

            throw new Exception("Euronext file download failed");
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
