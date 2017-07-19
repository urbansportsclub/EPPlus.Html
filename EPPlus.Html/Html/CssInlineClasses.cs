using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class CssInlineClasses : List<string>
    {
        public void Render(StringBuilder html)
        {
            html.Append("class=\"");
            html.Append
            (
                string.Join(" ", this
                    .Where(x => x != null))
            );
            html.Append('\"');
        }
    }
}
