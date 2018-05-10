using System.Collections.Generic;
using System.Text;

namespace EPPlus.Html.Html
{
    public class HtmlElementCollection : List<HtmlElement>, IRenderElement
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
