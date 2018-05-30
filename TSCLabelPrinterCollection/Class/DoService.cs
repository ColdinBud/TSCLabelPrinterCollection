using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSCLabelPrinterCollection.Class
{
    class DoService
    {
        public static string pubService = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString();
        public static string pubVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
    }
}
