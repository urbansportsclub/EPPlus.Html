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
            ExportExcel(ws, HtmlExportConfiguration.Minimal());
        }

        [TestMethod]
        public void OptionsAll()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, HtmlExportConfiguration.Default());
        }

        [TestMethod]
        public void OptionsBorders()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Borders = true });
        }

        [TestMethod]
        public void OptionsFill()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Fill = true });
        }

        [TestMethod]
        public void OptionsBordersAndFill()
        {
            ExcelWorksheet ws = GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Borders = true, Fill = true });
        }

        [TestMethod]
        public void OptionsNotSpecified()
        {
            ExcelWorksheet ws = GetWorkSheet();
            var html1 = ws.ToHtml(true);
            Assert.IsNotNull(html1, $"html generated without configuration is null");
            var html2 = ExportExcel(ws, HtmlExportConfiguration.Default());
            Assert.AreEqual(html1, html2, "HtmlExportConfiguration unspecified is not the same as HtmlExportConfiguration.Default()");
        }

        private string ExportExcel(ExcelWorksheet ws, HtmlExportConfiguration configuration)
        {
            var html = ws.ToHtml(configuration, true);
            Assert.IsNotNull(html, $"html generated with HtmlExportConfiguration {configuration.ToString()} is null");
            var outputFileName = OUTPUT_FILE.Replace("_", "_" + configuration.ToString());
            var cwd = Directory.GetCurrentDirectory();
            File.WriteAllText(Path.Combine(cwd, "../../TestOutput", outputFileName), html);
            return html;
        }
    }
}
