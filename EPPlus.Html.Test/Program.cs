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
            var package = new ExcelPackage(Test001);
            var worksheet = package.Workbook.Worksheets[1];

            string html = worksheet.ToHtml();

            string tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);

            Console.Write("Writing file to ");
            Console.WriteLine(tmpFile);
            Console.WriteLine("Press any key to exit");

            Console.ReadKey();
        }
    }
}
