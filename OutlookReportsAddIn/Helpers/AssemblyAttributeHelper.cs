using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn.Helpers
{
    public static class AssemblyAttributeHelper
    {
        public static string GetExecutingAssemblyAttribute<T>(Func<T, string> value) where T : Attribute
        {
            T attribute = (T)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(T));
            return value.Invoke(attribute);
        }
    }
}
