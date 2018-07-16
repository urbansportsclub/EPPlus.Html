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

            ProcessPictures(sheet);
            
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
                    htmlCell.Content = excelCell.Text;
                    htmlCell.Styles.Update(excelCell.ToCss());
                    
                    if (excelCell.Merge)
                    {
                        int mergeId = sheet.GetMergeCellId(row, col);
                        //skip next cells until no longer merged on same row
                        int mergeEndCol = col;
                        while (mergeEndCol <= lastCol && 
                               sheet.Cells[row, mergeEndCol].Merge &&
                               sheet.GetMergeCellId(row,mergeEndCol) == mergeId)
                        {
                            mergeEndCol++;
                        }

                        htmlCell.Attributes["colspan"] = (mergeEndCol - col).ToString();
                        htmlCell.Styles.Remove("width");
                        htmlCell.Styles.Remove("max-width");
                        col = mergeEndCol - 1;
                    }
                }
            }
            
            RemoveProcessedPictures(sheet);

            return htmlTable.ToString();
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
        
        private static void ProcessPictures(ExcelWorksheet workSheet)
        {
            //update all cells with image content
            var pictures = workSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>();

            foreach (var pic in pictures)
            {
                string imageSource;
                using (MemoryStream m = new MemoryStream())
                {
                    pic.Image.Save(m, System.Drawing.Imaging.ImageFormat.Png);

                    imageSource = string.Format("<img src=\"data:image/png;base64,{0}\"/>", Convert.ToBase64String(m.ToArray()));
                }

                workSheet.Cells[pic.From.Row + 1, pic.From.Column + 1].Value = imageSource;
            }
        }

        private static void RemoveProcessedPictures(ExcelWorksheet workSheet)
        {
            //update all cells with image content
            var pictures = workSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>();

            //remove image data
            foreach (var pic in pictures)
            {
                workSheet.Cells[pic.To.Row + 1, pic.To.Column + 1].Value = null;
            }
        }
    }
}
