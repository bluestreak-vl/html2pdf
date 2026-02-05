using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using PuppeteerSharp;
using PuppeteerSharp.Cdp;
using PdfSharp;
using System.Diagnostics;

class Program
{
    static async Task Main(string[] args)
    {
        // Define page range (edit for testing)
        int intStartPage = 1; // inclusive, set to 1 for first page or 0 to convert all
        int intEndPage = 0;   // inclusive, set to N for last page or 0 to convert all
        
        string strHtmlRootFolder;

        // Get root folder from args if provided. If no args, use current directory.
        if (args.Length == 0)
        {
            strHtmlRootFolder = Directory.GetCurrentDirectory();
            //check for "contents.db" in current directory
            if (!File.Exists(Path.Combine(strHtmlRootFolder, "contents.db")))
            {
                Console.WriteLine("Current directory does not contain contents.db.");
                Console.WriteLine("Please provide the HTML root folder, containing the contents.db file, as a command line argument.");
                // Show usage summary
                Console.WriteLine("Usage: HTML2PDF <html_root_folder> [start_page] [end_page]");
                Console.WriteLine("  html_root_folder: The root folder containing the contents.db file.");
                Console.WriteLine("  start_page: The first page to convert (inclusive, default: 1).");
                Console.WriteLine("  end_page: The last page to convert (inclusive, default: 0 = all pages).");
                return;
            }
        }
        else if (args.Length > 0 && Directory.Exists(args[0]))
        {
            strHtmlRootFolder = args[0];
            if (!File.Exists(Path.Combine(strHtmlRootFolder, "contents.db")))
            {
                Console.WriteLine($"The specified directory does not contain contents.db: {strHtmlRootFolder}");
                return;
            }
        }
        else
        {            
            Console.WriteLine("Please provide the HTML root folder, containing the contents.db file, as a command line argument.");
            return;
        }

        // second and third args can be startPage and endPage
        if (args.Length > 1 && int.TryParse(args[1], out int sp))
        {
            intStartPage = sp;
        }
        if (args.Length > 2 && int.TryParse(args[2], out int ep))
        {
            intEndPage = ep;
        }

        var contentsDbPath = Path.Combine(strHtmlRootFolder, "contents.db");

        if (!File.Exists(contentsDbPath))
        {
            Console.WriteLine($"Error: Database file not found: {contentsDbPath}");
            Debugger.Break();
            return;
        }

        // 1. Read the database
        using var db = new SqliteConnection($"Data Source={contentsDbPath}");
        db.Open();

        var pagesCmd = db.CreateCommand();
        pagesCmd.CommandText = "SELECT PID, Title, PageNum FROM Pages ORDER BY PageNum ASC";
        if (intStartPage > 0 && intEndPage > 0)
        {
            pagesCmd.CommandText = pagesCmd.CommandText.Replace("ORDER BY", "WHERE PageNum BETWEEN $startPage AND $endPage ORDER BY");
            pagesCmd.Parameters.AddWithValue("$startPage", intStartPage);
            pagesCmd.Parameters.AddWithValue("$endPage", intEndPage);
        }
        else if (intEndPage > 0)
        {
            pagesCmd.CommandText = pagesCmd.CommandText.Replace("ORDER BY", "WHERE PageNum <= $endPage ORDER BY");
            pagesCmd.Parameters.AddWithValue("$endPage", intEndPage);
        }
        else if (intStartPage > 1)
        {
            pagesCmd.CommandText = pagesCmd.CommandText.Replace("ORDER BY", "WHERE PageNum >= $startPage ORDER BY");
            pagesCmd.Parameters.AddWithValue("$startPage", intStartPage);
        }

        //Console.WriteLine($"Executing SQL: {pagesCmd.CommandText}");

        using var reader = pagesCmd.ExecuteReader();

        // Prepare Puppeteer
        var browserFetcher = new BrowserFetcher();
        var tskDownload = browserFetcher.DownloadAsync();
        // Wait up to 2 seconds for download. If the task is still running, notify user Chrome is being downloaded.
        if (await Task.WhenAny(tskDownload, Task.Delay(2000)) != tskDownload)
        {
            Console.WriteLine("Downloading Chromium browser. Please wait...");
            await tskDownload;
        }

        using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });

        Console.WriteLine("Starting conversion...");
        
        int intTickCountStart = System.Environment.TickCount;
        int intPageCount = 0;

        Stopwatch objStopwatch = new Stopwatch();
        objStopwatch.Start();

        // For each page in DB
        while (reader.Read())
        {
            var pid = reader.GetString(0);     // HTML file name
            var title = reader.GetString(1);   // Title
            var pageNum = reader.GetInt32(2);  // PageNum

            // Build HTML file path
            var htmlFilePath = Path.Combine(strHtmlRootFolder, $"{pid}.html");
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Warning: Not found {htmlFilePath}");
                Debugger.Break();
                continue;
            }

            // Sanitize Title for file system
            var safeTitle = Regex.Replace(title, @"[\\/:*?""<>|]", "_");
            //var strOutputPdf = Path.Combine(strHtmlRootFolder, $"P{pageNum.ToString("D3")}-{safeTitle}.pdf");
            //Save to a subfolder named "PDFs"
            var pdfOutputFolder = Path.Combine(strHtmlRootFolder, "PDFs");
            if (!Directory.Exists(pdfOutputFolder))
                Directory.CreateDirectory(pdfOutputFolder);
            var strOutputPdf = Path.Combine(pdfOutputFolder, $"P{pageNum.ToString("D3")}-{safeTitle}.pdf");

            Console.WriteLine($"Converting page {pageNum}: '{title}' ({pid}) to '{Path.GetFileName(strOutputPdf)}'");

            var fileUrl = new Uri(htmlFilePath).AbsoluteUri;

            using var pageLoader = await browser.NewPageAsync();
            await pageLoader.GoToAsync(fileUrl, WaitUntilNavigation.Networkidle0);

            // Get iframe if present
            CdpFrame iframePdf = null;
            if (pageLoader is CdpPage cdpPage)
            {
                foreach (var aframe in cdpPage.Frames)
                {
                    if (aframe.Name == "iframePdf" || aframe.Id == "iframePdf")
                    {
                        iframePdf = aframe as CdpFrame;
                        break;
                    }
                }
            }

            if (iframePdf == null)
            {
                Console.WriteLine("  Error: iframePdf not found");
                Debugger.Break();
                continue;
            }

            IPage pagePdf = pageLoader;
            if (iframePdf != null && !string.IsNullOrEmpty(iframePdf.Url))
            {
                pagePdf = await browser.NewPageAsync();
                await pagePdf.GoToAsync(iframePdf.Url, WaitUntilNavigation.Networkidle0);
                await Task.Delay(250);
            }

            // the page div is named pf[n], where n is a page number within the "chapter". 'n' is a hexidecimal number starting from 1.
            // eg. pf1, pf2 ..., pf9, pfa, pfb ... , pfd, pf10, pf11, etc.
            // Look for a div element with id that starts with "pf"

            var el = await pagePdf.QuerySelectorAsync("div[id^='pf']");
            if (el == null)
            {
                Console.WriteLine("  Skipped: #pf element not found.");
                Debugger.Break();
                continue;
            }

            var box = await el.BoundingBoxAsync();
            if (box == null)
            {
                Console.WriteLine("  Skipped: Could not get bounding box.");
                continue;
            }

            // You can adjust DPI here if needed
            const decimal dpi = 150.0M;
            var widthInInches = (box.Width / dpi).ToString("0.####", CultureInfo.InvariantCulture) + "in";
            var heightInInches = (box.Height / dpi).ToString("0.####", CultureInfo.InvariantCulture) + "in";

            await pagePdf.PdfAsync(strOutputPdf, new PdfOptions
            {
                PrintBackground = true,
                Width = widthInInches,
                Height = heightInInches,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions { Top = "0", Bottom = "0", Left = "0", Right = "0" },
                PreferCSSPageSize = false
            });

            // Remove PDF title
            using var pdfDoc = PdfSharp.Pdf.IO.PdfReader.Open(strOutputPdf, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
            pdfDoc.Info.Title = "";
            pdfDoc.Save(strOutputPdf);
            pdfDoc.Close();

            Console.WriteLine($"  Done: {Path.GetFileName(strOutputPdf)} [{box.Width}x{box.Height}px; {widthInInches} x {heightInInches}]");
            
            // Clean up page if opened as a new tab
            if (pagePdf != pageLoader)
                await pagePdf.CloseAsync();
            await pageLoader.CloseAsync();


            intPageCount++;
            if (objStopwatch.Elapsed.TotalMilliseconds >= 10000)
            {
                objStopwatch.Reset();
                objStopwatch.Start();
                GC.Collect();

                Console.WriteLine($"* * * Converted {intPageCount} pages... * * *");
            }
        }
        Console.WriteLine("All Done");
        Console.WriteLine($"Total time elapsed: {(System.Environment.TickCount - intTickCountStart) / 1000} seconds for {intPageCount} pages.");
    }
}