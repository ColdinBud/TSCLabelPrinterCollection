using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSCLabelPrinterCollection.Class;

namespace TSCLabelPrinterCollection.View
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
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
    }
}
