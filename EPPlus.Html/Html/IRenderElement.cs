using System.Text;

namespace EPPlus.Html.Html
{
    internal interface IRenderElement
    {
        void Render(StringBuilder html);
    }
}
