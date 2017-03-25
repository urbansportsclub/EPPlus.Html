using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    internal interface RenderElement
    {
        void Render(StringBuilder html);
    }
}
