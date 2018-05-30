using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TSCLabelPrinterCollection.Data;
using TSCLabelPrinterCollection.DLL;

namespace TSCLabelPrinterCollection.Class
{
    class TSCPrinter
    {
        string currentTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

        public void PrintOverTSC(string model, List<Label> labelList)
        {
            if (labelList.Count == 0)
            {
                return;
            }

            try
            {
                TSCLIB_DLL.openport(model);

                foreach (Label l in labelList)
                {
                    TSCLIB_DLL.setup("75", "30", "4", "8", "0", "2", "0");
                    TSCLIB_DLL.clearbuffer();
                    TSCLIB_DLL.printerfont("24", "24", "4", "0", "1", "1", l.Device);
                    TSCLIB_DLL.windowsfont(24, 60, 32, 0, 0, 0, "Arial", l.Drawer);
                    TSCLIB_DLL.windowsfont(336, 24, 32, 0, 0, 0, "Arial", currentTime);

                    if (l.MedName.Length <= 43)
                    {
                        TSCLIB_DLL.windowsfont(16, 96, 32, 0, 0, 0, "Arial", l.MedName);
                    }
                    else
                    {
                        TSCLIB_DLL.windowsfont(16, 96, 32, 0, 0, 0, "Arial", l.MedName.Substring(0, 43));
                        TSCLIB_DLL.windowsfont(16, 128, 32, 0, 0, 0, "Arial", l.MedName.Substring(43));
                    }

                    TSCLIB_DLL.windowsfont(24, 182, 48, 0, 0, 0, "標楷體", $"數量: {l.Amount}");
                    TSCLIB_DLL.barcode("264", "166", "128", "56", "1", "0", "2", "1", l.MedID);
                    TSCLIB_DLL.printlabel("1", "1");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                TSCLIB_DLL.closeport();
            }
        }
    }
}
