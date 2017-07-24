
using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using System;
using System.Linq;
using System.Text;

namespace EPPlus.Html
{
    public static class EPPlusExtensions 
    {
        public static string ToHtml(this ExcelWorksheet sheet)
        {
            return sheet.ToHtml(HtmlExportConfiguration.Default());
        }

        public static string ToHtml(this ExcelWorksheet sheet, bool consolidateStyles)
        {
            return sheet.ToHtml(HtmlExportConfiguration.Default(), consolidateStyles);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportConfiguration configuration)
        {
            return sheet.ToHtml(configuration, false);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportConfiguration configuration, bool consolidateStyles)
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
                if (configuration.Height)
                {
                    htmlRow.Styles.Update(excelRow.ToCss(configuration));
                }

                for (int col = 1; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    htmlCell.Content = excelCell.Text;
                    htmlCell.Styles.Update(excelCell.ToCss(configuration));
                }
            }

            // consolidate styles into classes
            if (consolidateStyles)
            {
                htmlTable.ClassTable = new CssClassTable();
                htmlTable.ConsolidateStyles(htmlTable.ClassTable);
            }

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
