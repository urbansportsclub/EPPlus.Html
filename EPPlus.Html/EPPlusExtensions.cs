using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html
{
    public static class EPPlusExtensions
    {
        public static string ToHtml(this ExcelWorksheet sheet)
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

                var test = excelRow.Style;
                HtmlElement htmlRow = htmlTable.AddChild("tr");
                htmlRow.Styles.Update(excelRow.ToCss());

                for (int col = 1; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    htmlCell.Content = excelCell.Value?.ToString();
                    htmlCell.Styles.Update(excelCell.ToCss());
                }
            }

            return htmlTable.ToString();
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
    }
}
