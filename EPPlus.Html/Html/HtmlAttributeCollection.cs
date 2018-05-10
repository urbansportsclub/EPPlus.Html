using System.Collections.Generic;
using System.Text;

namespace EPPlus.Html.Html
{
    public class HtmlAttributeCollection : Dictionary<string, object>, IRenderElement
    {
        public void Render(StringBuilder html)
        {
            foreach (var attr in this)
            {
                html.Append(attr.Key);
                if (attr.Value != null)
                {
                    html.Append("=\"");
                    html.Append(attr.Value.ToString());
                    html.Append("\" ");
                }
            }
        }
    }
}
