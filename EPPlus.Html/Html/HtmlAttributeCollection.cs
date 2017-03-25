using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class HtmlAttributeCollection : Dictionary<string, object>, RenderElement
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
