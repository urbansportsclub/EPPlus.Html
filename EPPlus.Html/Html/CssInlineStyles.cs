using System.Linq;
using System.Text;

namespace EPPlus.Html.Html
{
    public class CssInlineStyles : CssDeclaration, IRenderElement
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
