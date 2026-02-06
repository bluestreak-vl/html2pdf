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


    // To test with arguments in VS Code run from the terminal:  
    // dotnet run -- "C:\Path\To\HtmlRootFolder" 1 10

    private static bool _firstRun = true;
    private static decimal _dpi = 150.0M;

    static async Task Main(string[] args)
    {
        //For debugging
        //args = new string[] { @"D:\temp\ebooks\HTML_EBook", "6", "7" };

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
                Console.WriteLine($"The specified directory does not contain the contents.db: {strHtmlRootFolder}");
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

        
        SqliteDataReader rdrPages = await GetContentsDB(contentsDbPath, intStartPage, intEndPage);

        // Prepare Puppeteer
        var objBrowserFetcher = new BrowserFetcher();
        var tskDownload = objBrowserFetcher.DownloadAsync();
        // Wait up to 2 seconds for download. If the task is still running, notify user Chrome is being downloaded.
        if (await Task.WhenAny(tskDownload, Task.Delay(2000)) != tskDownload)
        {
            Console.WriteLine("Downloading Chromium browser. Please wait...");
            await tskDownload;
        }

        using var objBrowser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });

        Console.WriteLine("Starting conversion...");
        
        int intTickCountStart = System.Environment.TickCount;
        int intPageCount = 0;
        Stopwatch objStopwatch = new Stopwatch();
        objStopwatch.Start();

        // For each page in DB
        while (rdrPages.Read())
        {
            var pid = rdrPages.GetString(0);     // HTML file name
            var title = rdrPages.GetString(1);   // Title
            var pageNum = rdrPages.GetInt32(2);  // PageNum

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

            using var pageLoader = await objBrowser.NewPageAsync();
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
                pagePdf = await objBrowser.NewPageAsync();
                await pagePdf.GoToAsync(iframePdf.Url, WaitUntilNavigation.Networkidle0);
                await Task.Delay(250);
            }

            await ConvertPageToPdf(pagePdf, strOutputPdf);

            if (_firstRun && _dpi > 96)
            {
                Console.WriteLine("  Checking page count of first converted PDF to verify correct DPI setting...");
                await Task.Delay(250); 
                //check if converted PDF file has two pages. If so lower the dpi to 96 and run again

                using var pdfDocTest = PdfSharp.Pdf.IO.PdfReader.Open(strOutputPdf, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                if (pdfDocTest.PageCount > 1)   
                {
                    Console.WriteLine("  Notice: Converted PDF has more than one page. Adjusting DPI to 96 and reconverting.");
                    _dpi = 96.0M;
                    await ConvertPageToPdf(pagePdf, strOutputPdf);
                }

                _firstRun = false;
            }
            else
                _firstRun = false;
            

            // Clean up page if opened as a new tab
            if (pagePdf != pageLoader)
                await pagePdf.CloseAsync();
            await pageLoader.CloseAsync();


            intPageCount++;
            if (objStopwatch.Elapsed.TotalMilliseconds >= 10000)
            {
                objStopwatch.Restart();
                GC.Collect();

                Console.WriteLine($"* * * Converted {intPageCount} pages... * * *");
            }
        }
        
        rdrPages.Close();
        rdrPages.Dispose();

        Console.WriteLine("All Done");
        Console.WriteLine($"Total time elapsed: {(System.Environment.TickCount - intTickCountStart) / 1000} seconds for {intPageCount} pages.");
    }

    public static async Task<SqliteDataReader> GetContentsDB(string contentsDbPath, int startPage, int endPage)
    {
        // Read the database
        SqliteConnection db = new SqliteConnection($"Data Source={contentsDbPath}");
        db.Open();

        var cmdPages = db.CreateCommand();
        cmdPages.CommandText = "SELECT PID, Title, PageNum FROM Pages ORDER BY PageNum ASC";
        if (startPage > 0 && endPage > 0)
        {
            cmdPages.CommandText = cmdPages.CommandText.Replace("ORDER BY", "WHERE PageNum BETWEEN $startPage AND $endPage ORDER BY");
            cmdPages.Parameters.AddWithValue("$startPage", startPage);
            cmdPages.Parameters.AddWithValue("$endPage", endPage);
        }
        else if (endPage > 0)
        {
            cmdPages.CommandText = cmdPages.CommandText.Replace("ORDER BY", "WHERE PageNum <= $endPage ORDER BY");
            cmdPages.Parameters.AddWithValue("$endPage", endPage);
        }
        else if (startPage > 1)
        {
            cmdPages.CommandText = cmdPages.CommandText.Replace("ORDER BY", "WHERE PageNum >= $startPage ORDER BY");
            cmdPages.Parameters.AddWithValue("$startPage", startPage);
        }

        //Console.WriteLine($"Executing SQL: {pagesCmd.CommandText}");

        return cmdPages.ExecuteReader();
    }

    // Move PDF conversion code to seperate method
    public static async Task ConvertPageToPdf(IPage pagePdf, string outputPath)
    {

        // the page div is named pf[n], where n is a page number within the "chapter". 'n' is a hexidecimal number starting from 1.
        // eg. pf1, pf2 ..., pf9, pfa, pfb ... , pfd, pf10, pf11, etc.
        // Look for a div element with id that starts with "pf"

        var el = await pagePdf.QuerySelectorAsync("div[id^='pf']");
        if (el == null)
        {
            Console.WriteLine("  Error: #pf element not found.");
            Debugger.Break();
            return;
        }

        var box = await el.BoundingBoxAsync();
        if (box == null)
        {
            Console.WriteLine("  Error: Could not get bounding box.");
            Debugger.Break();
            return;
        }

        // You can adjust DPI here if needed
        
        var widthInInches = (box.Width / _dpi).ToString("0.####", CultureInfo.InvariantCulture) + "in";
        var heightInInches = (box.Height / _dpi).ToString("0.####", CultureInfo.InvariantCulture) + "in";

        try
        {
            await pagePdf.PdfAsync(outputPath, new PdfOptions
            {
                PrintBackground = true,
                Width = widthInInches,
                Height = heightInInches,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions { Top = "0", Bottom = "0", Left = "0", Right = "0" },
                PreferCSSPageSize = false
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error during PDF conversion: {ex.Message}");
            Debugger.Break();
            return;
        }

        // Remove PDF title
        using var pdfDoc = PdfSharp.Pdf.IO.PdfReader.Open(outputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        pdfDoc.Info.Title = "";
        pdfDoc.Save(outputPath);
        pdfDoc.Close();
        pdfDoc.Dispose();

        Console.WriteLine($"  Done: {Path.GetFileName(outputPath)} [{box.Width}x{box.Height}px; {widthInInches} x {heightInInches}]");
    }
}