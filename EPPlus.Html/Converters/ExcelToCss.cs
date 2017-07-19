using EPPlus.Html.Html;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Converters
{
    internal static class ExcelToCss
    {
        internal static CssDeclaration ToCss(this ExcelRow excelRow, HtmlExportOptions options)
        {
            var css = new CssDeclaration();

            if ((options & HtmlExportOptions.Height) == HtmlExportOptions.Height)
            {
                css["height"] = excelRow.Height + "px";
            }
            css.Update(excelRow.Style.ToCss(options));
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelRange excelRange)
        {
            return excelRange.ToCss(HtmlExportOptions.All);
        }

        internal static CssDeclaration ToCss(this ExcelRange excelRange, HtmlExportOptions options)
        {
            var css = new CssDeclaration();
            if (excelRange.Columns == 1 && excelRange.Rows == 1)
            {
                var excelColumn = excelRange.Worksheet.Column(excelRange.Start.Column);
                if ((options & HtmlExportOptions.Width) == HtmlExportOptions.Width)
                {
                    css["max-width"] = excelColumn.Width + "em";
                    css["width"] = excelColumn.Width + "em";
                }
                css.Update(excelRange.Style.ToCss(options));
            }
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle)
        {
            return excelStyle.ToCss(HtmlExportOptions.All);
        }

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle, HtmlExportOptions options)
        {
            var css = new CssDeclaration();
            if ((options & HtmlExportOptions.Fill) == HtmlExportOptions.Fill)
            {
                css["background-color"] = excelStyle.Fill.BackgroundColor.ToHexCode();
            }
            if ((options & HtmlExportOptions.Borders) == HtmlExportOptions.Borders)
            {
                css.Update(excelStyle.Border.ToCss());
            }
            if ((options & HtmlExportOptions.TextAlign) == HtmlExportOptions.TextAlign)
            {
                css["text-align"] = excelStyle.HorizontalAlignment.ToCssProperty();
            }
            css.Update(excelStyle.Font.ToCss(options));
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont)
        {
            return excelFont.ToCss(HtmlExportOptions.All);
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont, HtmlExportOptions options)
        {
            var css = new CssDeclaration();

            if (excelFont.Bold && (options & HtmlExportOptions.FontWeight) == HtmlExportOptions.FontWeight)
            {
                css["font-weight"] = "bold";
            }
            if ((options & HtmlExportOptions.FontFamily) == HtmlExportOptions.FontFamily)
            {
                css["font-family"] = excelFont.Name;
            }
            if ((options & HtmlExportOptions.FontSize) == HtmlExportOptions.FontSize)
            {
                css["font-size"] = excelFont.Size + "pt";
            }
            if ((options & HtmlExportOptions.FontColor) == HtmlExportOptions.FontColor)
            {
                css["color"] = excelFont.Color.ToHexCode();
            }
            return css;
        }

        internal static CssDeclaration ToCss(this Border border)
        {
            var css = new CssDeclaration();
            if (border.Top.Style != ExcelBorderStyle.None)
            {
                css["border-top"] = border.Top.ToCssProperty();
            }
            if (border.Bottom.Style != ExcelBorderStyle.None)
            {
                css["border-bottom"] = border.Bottom.ToCssProperty();
            }
            if (border.Right.Style != ExcelBorderStyle.None)
            {
                css["border-right"] = border.Right.ToCssProperty();
            }
            if (border.Left.Style != ExcelBorderStyle.None)
            {
                css["border-left"] = border.Left.ToCssProperty();
            };
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
                    return "1px solid";
                case ExcelBorderStyle.Thick:
                    return "2px solid";
                case ExcelBorderStyle.None:
                    return "none";
                default:
                    return "1px solid";
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
