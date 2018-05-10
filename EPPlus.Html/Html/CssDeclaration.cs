using System.Collections.Generic;
using System.Linq;

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
