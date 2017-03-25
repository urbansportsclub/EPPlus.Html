using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class HtmlElementCollection : List<HtmlElement>, RenderElement
    {
        public void Render(StringBuilder html)
        {
            foreach (HtmlElement element in this)
            {
                element.Render(html);
            }
        }

        public HtmlElement Add(string tagName)
        {
            HtmlElement element = new HtmlElement(tagName);
            this.Add(element);
            return element;
        }
    }
}
