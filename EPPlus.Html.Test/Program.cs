using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Test
{
    class Program
    {
        private static string CurrentLocation = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        private static FileInfo Test001 = new FileInfo(CurrentLocation + "\\Resources\\Test001.xlsx");

        static void Main(string[] args)
        {
            var package = new ExcelPackage(new FileInfo(CurrentLocation + "/Resources/Test001.xlsx"));
            var worksheet = package.Workbook.Worksheets[1];

            string html = worksheet.ToHtml();

            Show(html);
        }

        static void Show(string html)
        {
            string tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);
            Process.Start(tmpFile);
        }
    }
}
