using ServiceStack.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TSCLabelPrinterCollection.Data;

namespace TSCLabelPrinterCollection.Class
{
    class DoService
    {
        public static string pubService = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString();
        public static string pubVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        public static string PrintResult()
        {
            string result = "";
            string filePath = Properties.Settings.Default.FilePath;

            string dateTime = DateTime.Now.ToString("yyyyMMdd");
            string fileDateFormat = DateTime.Now.ToString("M-d-yyyy");

            Directory.CreateDirectory($"{filePath}\\{dateTime}");

            try
            {
                string[] sourceList = Directory.GetFiles(Properties.Settings.Default.SourcePath,
                    Properties.Settings.Default.FilePath + "*" + fileDateFormat + "*.csv", SearchOption.TopDirectoryOnly);

                foreach (string str in sourceList)
                {
                    string fileName = str.Substring(str.LastIndexOf("\\") + 1);

                    if (!File.Exists($"{filePath}\\{dateTime}\\{fileName}"))
                    {
                        if (!File.Exists($"{filePath}\\{fileName}"))
                        {
                            File.Copy(str, $"{filePath}\\{fileName}", false);
                        }
                    }
                }
            }
            catch
            {
                result += "連接無效，無權限存取網路空間\r\n";
            }
            finally
            {
                DirectoryInfo d = new DirectoryInfo(filePath);
                string[] csvList = Directory.GetFiles(filePath, "*.csv", SearchOption.TopDirectoryOnly);

                string time = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                if (csvList.Length > 0)
                {
                    FileInfo file = d.GetFiles("*.csv")[0];

                    var csv = File.ReadAllText(file.FullName);

                    result += file.FullName;

                    string[] propNames = null;
                    List<string[]> rows = new List<string[]>();
                    foreach (var line in CsvReader.ParseLines(csv))
                    {
                        string[] strArray = CsvReader.ParseFields(line).ToArray();
                        if (propNames == null)
                        {
                            propNames = strArray;
                        }
                        else
                        {
                            rows.Add(strArray);
                        }
                    }

                    List<Label> labelList = new List<Label>();
                    for (int r = 0; r < rows.Count; r++)
                    {
                        Label label = new Label();
                        var cells = rows[r];
                        for (int c = 0; c < cells.Length; c++)
                        {
                            switch (c)
                            {
                                case 0:
                                    label.Device = cells[c];
                                    break;
                                case 1:
                                    label.Drawer = cells[c];
                                    break;
                                case 2:
                                    label.MedID = cells[c];
                                    break;
                                case 3:
                                    label.MedName = cells[c];
                                    break;
                                case 12:
                                    label.Amount = cells[c];
                                    break;
                                default:
                                    break;
                            }
                        }
                        if (!(string.IsNullOrEmpty(label.MedID)))
                        {
                            labelList.Add(label);
                        }
                    }

                    result += $"{time} - File: {file.FullName}\r\n";
                    result += $"筆數: {labelList.Count}";

                    if (labelList.Count > 0)
                    {
                        string model = Properties.Settings.Default.TSCModel;
                        //new TSCPrinter().PrintOverTSC(model, labelList);
                    }

                    if (!File.Exists(filePath + "\\" + dateTime + "\\" + file.Name))
                    {
                        File.Move(file.FullName, filePath + "\\" + dateTime + "\\" + file.Name);
                    }
                    else
                    {
                        string nowTime = DateTime.Now.ToString("HHmm");
                        File.Move(file.FullName, filePath + "\\" + dateTime + "\\" + nowTime + file.Name);
                    }

                }
                else
                {
                    result += $"{time} - 沒有更新資料";
                }
            }

            return result;
        }
    }
}
