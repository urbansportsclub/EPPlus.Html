using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class CssInlineStyles : CssDeclaration, RenderElement
    {
        public void Render(StringBuilder html)
        {
            html.Append("style=\"");
            html.Append
            (
                string.Join("", this
                    .Where(x => x.Value != null)
                    .Select(x => x.Key + ":" + x.Value + ";"))
            );
            html.Append('\"');
        }
    }
}
