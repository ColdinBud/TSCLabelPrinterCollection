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
        public static string filePath = Properties.Settings.Default.FilePath;
        public static string dateTime = DateTime.Now.ToString("yyyyMMdd");

        public static string PrintResult()
        {
            string result = "";

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
                //DirectoryInfo d = new DirectoryInfo(filePath);
                string[] csvList = Directory.GetFiles(filePath, "*.csv", SearchOption.TopDirectoryOnly);

                result += ProcessCsvList(csvList);
            }

            return result;
        }

        public static string ProcessCsvList(string[] csvList, bool copy = false)
        {
            string result = "";
            string time = DateTime.Now.ToString("HH:mm:ss");

            int countCsv = 0;
            if (csvList.Length > 0)
            {
                //FileInfo file = d.GetFiles("*.csv")[0];

                //var csv = File.ReadAllText(file.FullName);
                //result += file.FullName;

                string fileName = csvList[countCsv].Substring(csvList[countCsv].LastIndexOf('\\') + 1);
                var csv = File.ReadAllText(csvList[countCsv]);
                result += $"資料處理中 - {csvList[countCsv]}\r\n";

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

                List<LabelData> labelList = new List<LabelData>();
                for (int r = 0; r < rows.Count; r++)
                {
                    LabelData label = new LabelData();
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

                result += $"{time}:    資料處理完畢 - {csvList[countCsv]}\r\n";
                result += $"筆數: {labelList.Count}\r\n";

                if (labelList.Count > 0)
                {
                    string model = Properties.Settings.Default.TSCModel;
                    new TSCPrinter().PrintOverTSC(model, labelList);
                }

                if (!File.Exists(filePath + "\\" + dateTime + "\\" + fileName))
                {
                    if (copy)
                    {
                        File.Copy(csvList[countCsv], filePath + "\\" + dateTime + "\\" + fileName);
                    }
                    else
                    {
                        File.Move(csvList[countCsv], filePath + "\\" + dateTime + "\\" + fileName);
                    }
                }
                else
                {
                    string nowTime = DateTime.Now.ToString("HHmm");
                    if (copy)
                    {
                        File.Copy(csvList[countCsv], filePath + "\\" + dateTime + "\\" + nowTime + fileName);
                    }
                    else
                    {
                        File.Move(csvList[countCsv], filePath + "\\" + dateTime + "\\" + nowTime + fileName);
                    }
                }
                countCsv++;
            }
            else
            {
                result += $"沒有更新資料\r\n";
            }

            return result;
        }
    }
}
