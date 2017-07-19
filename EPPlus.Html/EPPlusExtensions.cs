using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using System;
using System.Linq;
using System.Text;

namespace EPPlus.Html
{

#pragma warning disable S101 // Types should be named in camel case
    public static class EPPlusExtensions 
#pragma warning restore S101 // Types should be named in camel case
    {
        public static string ToHtml(this ExcelWorksheet sheet)
        {
            return sheet.ToHtml(HtmlExportOptions.All);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportOptions options)
        {
            return sheet.ToHtml(options, false);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportOptions options, bool consolidateStyles)
        {
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;

            HtmlElement htmlTable = new HtmlElement("table");
            htmlTable.Attributes["cellspacing"] = 0;
            htmlTable.Styles["white-space"] = "nowrap";

            //render rows
            for (int row = 1; row <= lastRow; row++)
            {
                ExcelRow excelRow = sheet.Row(row);

                HtmlElement htmlRow = htmlTable.AddChild("tr");
                if ((options & HtmlExportOptions.Height) == HtmlExportOptions.Height)
                {
                    htmlRow.Styles.Update(excelRow.ToCss(options));
                }

                for (int col = 1; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    htmlCell.Content = excelCell.Text;
                    htmlCell.Styles.Update(excelCell.ToCss(options));
                }
            }

            // consolidate styles into classes
            htmlTable.ClassTable = new CssClassTable();
            htmlTable.ConsolidateStyles(htmlTable.ClassTable);

            return htmlTable.ToString();
        }

        public static void ConsolidateStyles(this HtmlElement elmt, CssClassTable classTable)
        {
            if (elmt.Styles.Any())
            {
                StringBuilder sb = new StringBuilder();
                foreach (var style in elmt.Styles)
                {
                    sb.Append($"{style.Key}:{style.Value};");
                }

                var styleDef = sb.ToString();
                var className = classTable.GetOrAddClassForStyle(styleDef);
                elmt.InlineClasses.Add(className);
            }
            foreach (var child in elmt.Children)
            {
                ConsolidateStyles(child, classTable);
            }
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
    }
}
