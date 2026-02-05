# HTML2PDF
A C# console application that converts individual pages from PDF2HTML output to separate PDF files using PuppeteerSharp.

Basically, it creates PDFs from the output of https://github.com/shebinleo/pdf2html

## Usage

HTML2PDF.exe <html_root_folder> [start_page] [end_page]
- html_root_folder: The root folder containing the contents.db file.
- start_page: The first page to convert (inclusive, default: 1).
- end_page: The last page to convert (inclusive, default: 0 = all pages).
