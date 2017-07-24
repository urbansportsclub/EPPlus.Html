using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using System.Text;
using EPPlus.Html;


namespace Test
{
    [TestClass]
    public class BasicTests
    {
        private static readonly string INPUT_FILE = "test.xlsx";
        private static readonly string OUTPUT_FILE = "test_.html";

        private ExcelWorksheet GetWorkSheet()
        {
            var cwd = Directory.GetCurrentDirectory();
            FileInfo file = new FileInfo(Path.Combine(cwd, INPUT_FILE));
            ExcelPackage xlPackage = new ExcelPackage(file);
            if (xlPackage.Workbook == null)
            {
                Console.WriteLine("Excel package has no workbook");
                return null;
            }

            if (xlPackage.Workbook.Worksheets == null)
            {
                Console.WriteLine("Excel package has workbook without worksheets");
                return null;
            }
            Console.WriteLine($"Workbook has {xlPackage.Workbook.Worksheets.Count} worksheets");
            if (xlPackage.Workbook.Worksheets.Count < 1)
            {
                Console.WriteLine("Excel package has workbook without worksheets");
                return null;
            }
            return xlPackage.Workbook.Worksheets[1];
        }

        [TestMethod]
        public void OptionsNone()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, HtmlExportOptions.None);
        }

        [TestMethod]
        public void OptionsAll()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, HtmlExportOptions.All);
        }

        [TestMethod]
        public void OptionsBorders()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, HtmlExportOptions.Borders);
        }

        [TestMethod]
        public void OptionsFill()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, HtmlExportOptions.Fill);
        }

        [TestMethod]
        public void OptionsBordersAndFill()
        {
            ExcelWorksheet ws = GetWorkSheet();
            var html1 = ExportExcel(ws, HtmlExportOptions.BordersAndFill);
            var html2 = ExportExcel(ws, HtmlExportOptions.Borders | HtmlExportOptions.Fill);
            Assert.AreEqual(html1, html2, "Combined HtmlExportOptions 'BordersAndFill' do not equal the merged options");
        }

        [TestMethod]
        public void OptionsNotSpecified()
        {
            ExcelWorksheet ws = GetWorkSheet();
            var html1 = ws.ToHtml(true);
            Assert.IsNotNull(html1, $"html generated without is null");
            var html2 = ExportExcel(ws, HtmlExportOptions.All);
            Assert.AreEqual(html1, html2, "HtmlExportOptions unspecified is not the same as HtmlExportOptions.All");
        }

        private string ExportExcel(ExcelWorksheet ws, HtmlExportOptions options)
        {
            var html = ws.ToHtml(options, true);
            Assert.IsNotNull(html, $"html generated with HtmlExportOption {options.ToString()} is null");
            var outputFileName = OUTPUT_FILE.Replace("_", "_" + options.ToString());
            var cwd = Directory.GetCurrentDirectory();
            File.WriteAllText(Path.Combine(cwd, "../../TestOutput", outputFileName), html);
            return html;
        }
    }
}
