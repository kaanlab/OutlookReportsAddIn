using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn.ViewModels
{
    public class AboutViewModel : BaseViewModel
    {
        public static string AssemblyVersion { get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); } }
        public static string AssemblyCopyright { get { return GetExecutingAssemblyAttribute<AssemblyCopyrightAttribute>(a => a.Copyright); } }
        public static string AssemblyCompany { get { return GetExecutingAssemblyAttribute<AssemblyCompanyAttribute>(a => a.Company); } }
        public static string WindowTitle { get => "Об этом дополнении"; }

        private static string GetExecutingAssemblyAttribute<T>(Func<T, string> value) where T : Attribute
        {
            T attribute = (T)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(T));
            return value.Invoke(attribute);
        }
    }
}
