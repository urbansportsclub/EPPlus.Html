using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class HtmlElement : RenderElement
    {
        public string TagName { get; private set; }
        public string Content { get; set; }
        public CssInlineStyles Styles { get; private set; }
        public HtmlAttributeCollection Attributes { get; private set; }
        public HtmlElementCollection Children { get; private set; }

        public HtmlElement(string tagName)
        {
            this.TagName = tagName;
            this.Styles = new CssInlineStyles();
            this.Attributes = new HtmlAttributeCollection();
            this.Children = new HtmlElementCollection();
        }

        public HtmlElement AddChild(HtmlElement element)
        {
            this.Children.Add(element);
            return element;
        }

        public HtmlElement AddChild(string tagName)
        {
            return this.Children.Add(tagName);
        }

        public void Render(StringBuilder html)
        {
            html.Append("<");
            html.Append(this.TagName);

            if (this.Styles.Any())
            {
                html.Append(" ");
                this.Styles.Render(html);
                html.Append(" ");
            }

            if (this.Attributes.Any())
            {
                html.Append(" ");
                this.Attributes.Render(html);
            }

            html.Append(">");

            if (this.Content != null)
            {
                html.Append(this.Content);
            }

            this.Children.Render(html);
            html.Append("</");
            html.Append(this.TagName);
            html.Append(">");
        }

        public override string ToString()
        {
            var html = new StringBuilder();
            this.Render(html);
            return html.ToString();
        }
    }
}
