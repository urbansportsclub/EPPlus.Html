using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Html
{
    public static class EPPlusExtensions
    {
        public static string ToHtml(this ExcelWorksheet sheet, ExcelAddress exportAddressRange = null)
        {
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;
            int startRow = 1;
            int startColumn = 1;
            //cache what cells we have processed on the merge
            List<(int row, int col)> cache = new List<(int, int)>();

            if (exportAddressRange != null)
            {
                startColumn = exportAddressRange.Start.Column;
                startRow = exportAddressRange.Start.Row;
                lastCol = exportAddressRange.End.Column;
                lastRow = exportAddressRange.End.Row;
            }

            HtmlElement htmlTable = new HtmlElement("table");

            htmlTable.Attributes["cellspacing"] = 0;
            htmlTable.Styles["white-space"] = "nowrap";
            htmlTable.Styles["table-layout"] = "auto";

            //render rows
            for (int row = startRow; row <= lastRow; row++)
            {
                ExcelRow excelRow = sheet.Row(row);

                HtmlElement htmlRow = htmlTable.AddChild("tr");

                htmlRow.Styles.Update(excelRow.ToCss());

                for (int col = startColumn; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = null;

                    if (!cache.Any(c => c.row == row && c.col == col))
                    { 
                        htmlCell = htmlRow.AddChild("td");
                    }

                //once we find a cell thats merged, we will have to iterate through the cells until we find the end
                    if (excelCell.Merge && !cache.Any(c=> c.row == row && c.col == col))
                    {
                        var currentCellColumn = excelCell.Columns;
                        var currentCellRow = excelCell.Rows;
                        int mergeColumnSize = 1;
                        int mergeRowSize = 1;
                        int lastMergedColumn = currentCellColumn;

                        cache.Add((currentCellRow, currentCellColumn));

                        //start at the next column and keep going until we find the last merged column
                        for (int i = currentCellColumn+1; i < sheet.Dimension.Columns+1; i++)
                        {
                            if (sheet.Cells[excelCell.Rows, i].Merge)
                            {
                                mergeColumnSize++;
                                lastMergedColumn = i;
                                cache.Add((excelCell.Rows, i));
                            }

                            else
                            {
                                //merge has ended. leave the loop
                                break;
                            }
                        }

                        //since the merged cells are always a box shape, start at the column were the merge ended and go down the rows until we hit the end of the merge
                        for (int i = currentCellRow+1; i < sheet.Dimension.Rows+1; i++)
                        {
                            if (sheet.Cells[i, lastMergedColumn].Merge)
                            {
                                mergeRowSize++;

                                //add all the merged columns in the row
                                for (int j = excelCell.Columns; j < lastMergedColumn; j++)
                                {
                                    cache.Add((i, j));
                                }
                            }

                            else
                            {                               
                                //merge has ended. leave the loop
                                break;
                            }
                        }

                        if (mergeColumnSize > 1)
                        {
                            htmlCell.Attributes["colspan"] = mergeColumnSize.ToString();
                        }

                        if (mergeRowSize > 1)
                        {
                            htmlCell.Attributes["rowspan"] = mergeRowSize.ToString();
                        }

                        //skip the merged cells since they are technically empty cells


                    }

                    if (htmlCell != null)
                    {
                        htmlCell.Content = excelCell.Text;
                        htmlCell.Styles.Update(excelCell.ToCss());
                    }
                }
            }

            return htmlTable.ToString();
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = ((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
    }
}
