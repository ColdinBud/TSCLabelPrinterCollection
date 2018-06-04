using ServiceStack.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSCLabelPrinterCollection.Class;
using TSCLabelPrinterCollection.Data;

namespace TSCLabelPrinterCollection.View
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        List<FormularyData> formularyList = new List<FormularyData>();
        ArrayList deviceList = new ArrayList();

        private void FormMain_Load(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            var csv = File.ReadAllText($"{path}\\Formulary Report.csv");
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

            for (int r = 0; r < rows.Count; r++)
            {
                FormularyData formularyData = new FormularyData();
                var cells = rows[r];
                for (int c = 0; c < cells.Length; c++)
                {
                    switch (c)
                    {
                        case 0:
                            formularyData.GenericName = cells[c];
                            break;
                        case 4:
                            formularyData.MedName = cells[c];
                            break;
                        case 5:
                            formularyData.MedID = cells[c];
                            break;
                        case 14:
                            formularyData.DosageForm = cells[c];
                            break;
                        case 16:
                            formularyData.StrengthUnit = cells[c];
                            break;
                        case 17:
                            formularyData.ConcentrationVolumeUnit = cells[c];
                            break;
                        case 18:
                            formularyData.TotalVolumeUnit = cells[c];
                            break;
                        default:
                            break;
                    }
                }
                formularyList.Add(formularyData);
            }

            var deviceCsv = File.ReadAllText($"{path}\\Device List.csv");
            string[] devicePropNames = null;
            List<string[]> deviceRows = new List<string[]>();
            foreach (var line in CsvReader.ParseLines(deviceCsv))
            {
                string[] strArray = CsvReader.ParseFields(line).ToArray();
                if (devicePropNames == null)
                {
                    devicePropNames = strArray;
                }
                else
                {
                    deviceRows.Add(strArray);
                }
            }

            for (int r = 0; r < deviceRows.Count; r++)
            {
                deviceList.Add(deviceRows[r][0]);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (richTextBox1.Lines.Length > 14)
            {
                richTextBox1.Clear();
            }
            richTextBox1.Text += string.IsNullOrEmpty(richTextBox1.Text) ? "\r\n" : "";

            string printResult = DoService.PrintResult();
            string res = string.IsNullOrEmpty(printResult) ? "沒有新資料" : printResult;

            richTextBox1.Text += $"{DateTime.Now.ToString("HH:mm:ss")}:    {res}\r\n"; ;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormConfig objConfig = new FormConfig();
            if (objConfig.ShowDialog(this) == DialogResult.OK)
            {
            }
            objConfig.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Interval = Properties.Settings.Default.Interval * 1000;
            timer1.Enabled = true;
            button1.Enabled = false;
            button3.Enabled = true;

            richTextBox1.Text += "排程已啟動.....\r\n";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            button1.Enabled = true;
            button3.Enabled = false;
        }

        List<Label> deviceLabelList = new List<Label>();
        List<Label> medIDLabelList = new List<Label>();
        List<Label> amountLabelList = new List<Label>();
        List<ComboBox> deviceComboBoxList = new List<ComboBox>();
        List<TextBox> medIDTextBoxList = new List<TextBox>();
        List<TextBox> amountTextBoxList = new List<TextBox>();
        List<Label> medNameLabelList = new List<Label>();

        private void button4_Click(object sender, EventArgs e)
        {
            deviceLabelList.Add(new Label());
            medIDLabelList.Add(new Label());
            amountLabelList.Add(new Label());
            deviceComboBoxList.Add(new ComboBox());
            medIDTextBoxList.Add(new TextBox());
            amountTextBoxList.Add(new TextBox());
            medNameLabelList.Add(new Label());

            string[] arrString = (string[])deviceList.ToArray(typeof(string));

            for (int l = 0; l < medIDLabelList.Count; l++)
            {
                deviceLabelList[l].Text = "裝置名稱:";
                deviceLabelList[l].Bounds = new Rectangle(10, (l * 40), 80, 40);

                deviceComboBoxList[l].Items.Clear();
                deviceComboBoxList[l].Items.AddRange(arrString);
                deviceComboBoxList[l].Bounds = new Rectangle(90, (l * 40), 70, 40);

                medIDLabelList[l].Text = "藥品八碼:";
                medIDLabelList[l].Bounds = new Rectangle(180, (l * 40), 80, 40);

                medIDTextBoxList[l].Name = "TextBox" + l;
                medIDTextBoxList[l].Bounds = new Rectangle(260, (l * 40), 90, 40);
                medIDTextBoxList[l].TextChanged += new EventHandler(t_TextChanged);

                amountLabelList[l].Text = "數量:";
                amountLabelList[l].Bounds = new Rectangle(360, (l * 40), 50, 40);

                amountTextBoxList[l].Bounds = new Rectangle(410, (l * 40), 50, 40);

                //medNameLabelList[l].Text = "";
                medNameLabelList[l].Bounds = new Rectangle(480, (l * 40), 440, 40);

                panel2.Controls.Add(deviceLabelList[l]);
                panel2.Controls.Add(deviceComboBoxList[l]);
                panel2.Controls.Add(medIDLabelList[l]);
                panel2.Controls.Add(medIDTextBoxList[l]);
                panel2.Controls.Add(amountLabelList[l]);
                panel2.Controls.Add(amountTextBoxList[l]);
                panel2.Controls.Add(medNameLabelList[l]);
            }
        }

        void t_TextChanged(object sender, EventArgs e)
        {
            var msg = formularyList.FirstOrDefault(x => x.MedID.Equals(((TextBox)sender).Text.ToUpper()));

            int name = 0;
            int.TryParse(((TextBox)sender).Name.Substring(7), out name);

            if (msg != null)
            {
                string total = String.IsNullOrEmpty(msg.TotalVolumeUnit) ? "" : $"({msg.TotalVolumeUnit})";

                medNameLabelList[name].Text = $"{msg.GenericName} ({msg.MedName}) {msg.StrengthUnit} " +
                    $"{msg.ConcentrationVolumeUnit} {total} {msg.DosageForm}";
                panel2.Controls.Add(medNameLabelList[name]);
            }
            else if (medNameLabelList[name].Text != "")
            {
                medNameLabelList[name].Text = "";
                panel2.Controls.Add(medNameLabelList[name]);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel2.Controls.Remove(deviceLabelList[deviceLabelList.Count - 1]);
            panel2.Controls.Remove(deviceComboBoxList[deviceComboBoxList.Count - 1]);
            panel2.Controls.Remove(medIDLabelList[medIDLabelList.Count - 1]);
            panel2.Controls.Remove(medIDTextBoxList[medIDTextBoxList.Count - 1]);
            panel2.Controls.Remove(amountLabelList[amountLabelList.Count - 1]);
            panel2.Controls.Remove(amountTextBoxList[amountTextBoxList.Count - 1]);
            panel2.Controls.Remove(medNameLabelList[medNameLabelList.Count - 1]);

            deviceLabelList.RemoveAt(deviceLabelList.Count - 1);
            deviceComboBoxList.RemoveAt(deviceComboBoxList.Count - 1);
            medIDLabelList.RemoveAt(medIDLabelList.Count - 1);
            medIDTextBoxList.RemoveAt(medIDTextBoxList.Count - 1);
            amountLabelList.RemoveAt(amountLabelList.Count - 1);
            amountTextBoxList.RemoveAt(amountTextBoxList.Count - 1);
            medNameLabelList.RemoveAt(medNameLabelList.Count - 1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            List<LabelData> printLabelList = new List<LabelData>();
            for (int i = 0; i < medIDLabelList.Count; i++)
            {
                string device = deviceComboBoxList[i].Text;
                string medID = medIDTextBoxList[i].Text.ToUpper();
                string amount = amountTextBoxList[i].Text;
                string medName = medNameLabelList[i].Text;

                if (!(string.IsNullOrEmpty(device) || string.IsNullOrEmpty(medID) ||
                    string.IsNullOrEmpty(amount) || string.IsNullOrEmpty(medName)))
                {
                    printLabelList.Add(new LabelData { Device = device, MedID = medID, MedName = medName, Amount = amount });
                }
            }

            if (printLabelList.Count > 0)
            {
                new TSCPrinter().PrintOverTSC(Properties.Settings.Default.TSCModel, printLabelList);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Title = "請選擇 Station 檔案";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "csv files (*.*)|*.csv";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string[] csvList = { dialog.FileName };
                richTextBox1.Text += DoService.ProcessCsvList(csvList, true);
            }
        }
    }
}
