using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html
{
    public class HtmlExportConfiguration
    {
        public bool TextAlign { get; set; }
        public bool Fill { get; set; }
        public bool Borders { get; set; }
        public bool Width { get; set; }
        public bool Height { get; set; }
        public bool FontFamily { get; set; }
        public bool FontSize { get; set; }
        public bool FontWeight { get; set; }
        public bool FontColor { get; set; }

        public static HtmlExportConfiguration Default()
        {
            return new HtmlExportConfiguration()
            {
                TextAlign = true,
                Fill = true,
                Width = true,
                Height = true,
                FontFamily = true,
                FontSize = true,
                FontWeight = true,
                FontColor = true
            };
        }

        public static HtmlExportConfiguration Minimal()
        {
            return new HtmlExportConfiguration();
        }
    }
}
