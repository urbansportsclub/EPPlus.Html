using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class CssDeclaration : Dictionary<string, string>
    {
        public void Update(CssDeclaration declarations)
        {
            foreach (var declaration in declarations.Where(x => x.Value != null))
            {
                this[declaration.Key] = declaration.Value;
            }
        }
    }
}
