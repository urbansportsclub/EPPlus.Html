using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlus.Html.Test
{
    [TestClass]
    public class HtmlTests
    {
        private readonly string CurrentLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        [TestMethod]
        public void ExcelToHtml()
        {
            FileInfo test001 = new FileInfo(@"Resources/Test001.xlsx");
            var package = new ExcelPackage(test001);
            var worksheet = package.Workbook.Worksheets[1];

            var html = worksheet.ToHtml();

            Show(html);
        }
        //
        [TestMethod]
        public void ExcelToHtml1()
        {
            FileInfo test001 = new FileInfo(@"C:\Users\adkerti\Desktop\2018 KL Auto Dim PROXI As Of 05-07-2018 9.30 AM (8).xlsx");
            var package = new ExcelPackage(test001);
            var worksheet = package.Workbook.Worksheets[1];

            var html = worksheet.ToHtml(new ExcelAddress("A1:B14"));

            Show(html);
        }

        [TestMethod]
        public void ExcelToHtml2()
        {
            FileInfo test001 = new FileInfo(@"\\WIN-SERVER\Public\DB-FlashReports-Test\FlashActivities\Projects\2018 DS Harman Radio Validate v17.25.03\DailyLogs\2018 DS Harman Radio Validate v17.25.03 As Of 04-09-2018 9.19 AM (16).xlsx");
            var package = new ExcelPackage(test001);
            var worksheet = package.Workbook.Worksheets[1];

            var html = worksheet.ToHtml();

            Show(html);
        }

        private void Show(string html)
        {
            var tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);
            Process.Start(tmpFile);
        }
    }
}