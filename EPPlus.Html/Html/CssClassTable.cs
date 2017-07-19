using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Html
{
    public class CssClassTable : Dictionary<string, string>
    {
        int classCounter = 1;
        public string GetOrAddClassForStyle(string styleDefinition)
        {
            if (!ContainsKey(styleDefinition))
            {
                var className = "table-style-" + classCounter++;
                this.Add(styleDefinition, className);
            }
            return this[styleDefinition];
        }
    }
}
