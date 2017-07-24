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
        internal static CssDeclaration ToCss(this ExcelRow excelRow, HtmlExportConfiguration configuration)
        {
            var css = new CssDeclaration();

            if (configuration.Height)
            {
                css["height"] = excelRow.Height + "px";
            }
            css.Update(excelRow.Style.ToCss(configuration));
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelRangeBase excelRange)
        {
            return excelRange.ToCss(HtmlExportConfiguration.Default());
        }

        internal static CssDeclaration ToCss(this ExcelRangeBase excelRange, HtmlExportConfiguration configuration)
        {
            var css = new CssDeclaration();
            if (excelRange.Columns == 1 && excelRange.Rows == 1)
            {
                var excelColumn = excelRange.Worksheet.Column(excelRange.Start.Column);
                if (configuration.Width)
                {
                    css["max-width"] = excelColumn.Width + "em";
                    css["width"] = excelColumn.Width + "em";
                }
                css.Update(excelRange.Style.ToCss(configuration));
            }
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle)
        {
            return excelStyle.ToCss(HtmlExportConfiguration.Default());
        }

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle, HtmlExportConfiguration configuration)
        {
            var css = new CssDeclaration();
            if (configuration.Fill)
            {
                css["background-color"] = excelStyle.Fill.BackgroundColor.ToHexCode();
            }
            if (configuration.Borders)
            {
                css.Update(excelStyle.Border.ToCss());
            }
            if (configuration.TextAlign)
            {
                css["text-align"] = excelStyle.HorizontalAlignment.ToCssProperty();
            }
            css.Update(excelStyle.Font.ToCss(configuration));
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont)
        {
            return excelFont.ToCss(HtmlExportConfiguration.Default());
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont, HtmlExportConfiguration configuration)
        {
            var css = new CssDeclaration();

            if (excelFont.Bold && configuration.FontWeight)
            {
                css["font-weight"] = "bold";
            }
            if (configuration.FontFamily)
            {
                css["font-family"] = excelFont.Name;
            }
            if (configuration.FontSize)
            {
                css["font-size"] = excelFont.Size + "pt";
            }
            if (configuration.FontColor)
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
