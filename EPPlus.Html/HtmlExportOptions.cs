using System;

namespace EPPlus.Html
{
    [Flags]
    public enum HtmlExportOptions
    {
        None = 0,
        TextAlign = 1,
        Borders = 2,
        Fill = 4,
        BordersAndFill = 6, // combination
        Width = 8,
        Height = 16,
        WidthAndHeight = 24, // combination
        FontFamily = 32,
        FontSize = 64,
        FontWeight = 128,
        FontColor = 256,
        Font = 480, // combination
        All = 511 // combination
    }
}