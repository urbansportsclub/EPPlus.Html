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
        public void ExcelToHtml1()
        {
            FileInfo test001 = new FileInfo(CurrentLocation + "\\Resources\\Test001.xlsx");
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