using EPPlus.Html.Html;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Html.Converters
{
    internal static class ExcelToCss
    {
        internal static CssDeclaration ToCss(this ExcelRow excelRow)
        {
            var css = new CssDeclaration();

            css["height"] = excelRow.Height + "px";

            css.Update(excelRow.Style.ToCss(false, false, false));

            return css;
        }

        internal static CssDeclaration ToCss(this ExcelRange excelRange)
        {
            var css = new CssDeclaration();
            if (excelRange.Columns == 1 && excelRange.Rows == 1)
            {
                //var excelColumn = excelRange.Worksheet.Column(excelRange.Start.Column);

                //the tables ignore this width anyways
                //css["max-width"] = excelColumn.Width + "em";
                //css["width"] = excelColumn.Width + "em";
                css["overflow"] = excelRange.Worksheet.Cells[excelRange.End.Row, excelRange.End.Column + 1].Value != null ? "hidden" : null;

                css.Update(excelRange.Style.ToCss(excelRange.Worksheet.Dimension.Rows == excelRange.End.Row, excelRange.Worksheet.Dimension.Columns == excelRange.End.Column, excelRange.Merge));
            }
            return css;
        }      

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle, bool isLastRow, bool isLastColumn, bool isMerged)
        {
            var css = new CssDeclaration();
            css["text-align"] = excelStyle.HorizontalAlignment.ToCssProperty();
            css["background-color"] = excelStyle.Fill.BackgroundColor.ToHexCode();
            css["overflow"] = excelStyle.HorizontalAlignment == ExcelHorizontalAlignment.Fill ? "hidden" : null;
            css["padding-left"] = "5px";
            css["padding-right"] = "5px";

            if (excelStyle.Indent > 0)
            {
                css["padding-left"] = $"{excelStyle.Indent * 10}px";
            }

            css.Update(excelStyle.Font.ToCss());
            css.Update(excelStyle.Border.ToCss(isLastRow, isLastColumn, isMerged));
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont)
        {
            var css = new CssDeclaration();

            if (excelFont.Bold)
            {
                css["font-weight"] = "bold";
            }

            css["font-family"] = excelFont.Name;
            css["font-size"] = excelFont.Size + "pt";

            css["color"] = excelFont.Color.ToHexCode();

            return css;
        }

        internal static CssDeclaration ToCss(this Border border, bool isLastRow, bool isLastColumn, bool isMerged)
        {
            var css = new CssDeclaration();

            css["border-top"] = border.Top.ToCssProperty();

            css["border-left"] = border.Left.ToCssProperty();

            //prevent double borders
            if (isLastRow)
            {
                css["border-bottom"] = border.Bottom.ToCssProperty();
            }

            //prevent double borders
            if (isLastColumn || isMerged)
            {
                css["border-right"] = border.Right.ToCssProperty();
            }

            return css;
        }

        internal static string ToCssProperty(this ExcelBorderItem excelBorderItem)
        {
            if (excelBorderItem != null)
            {
                string color = (excelBorderItem.Color.Rgb != null)
                    ? excelBorderItem.Color.ToHexCode()
                    : "black";

                return excelBorderItem.Style.ToCssProperty() + " " + color;
            }
            else
            {
                return null;
            }
        }

        internal static string ToCssProperty(this ExcelBorderStyle excelBorderStyle)
        {
            switch (excelBorderStyle)
            {
                case ExcelBorderStyle.Thin:
                    return ".5px solid";
                case ExcelBorderStyle.Medium:
                    return "1px solid";
                case ExcelBorderStyle.Thick:
                    return "2px solid";
                default:
                    return "none";
            }
        }

        internal static string ToCssProperty(this ExcelHorizontalAlignment excelHorizontalAlignment)
        {
            switch (excelHorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Left:
                    return "left";
                case ExcelHorizontalAlignment.Center:
                    return "center";
                default:
                    return null;
            }
        }

        internal static string ToHexCode(this ExcelColor excelColor)
        {
            if (excelColor != null && excelColor.Rgb != null && excelColor.Rgb.Length > 3)
            {
                return "#" + excelColor.Rgb.Substring(2);
            }
            else
            {
                return null;
            }
        }
    }
}
