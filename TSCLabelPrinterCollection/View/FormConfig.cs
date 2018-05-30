using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSCLabelPrinterCollection.Class;

namespace TSCLabelPrinterCollection.View
{
    public partial class FormConfig : Form
    {
        public DataTable dtConfig = CreateDataTable();

        public FormConfig()
        {
            InitializeComponent();
        }

        private static DataTable CreateDataTable()
        {
            DataTable cfgTable = new DataTable();
            cfgTable.Columns.Add("columnTitle", typeof(String));
            cfgTable.Columns.Add("columnValue", typeof(String));

            return cfgTable;
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            Console.WriteLine("COUNT= " + Properties.Settings.Default.Properties.Count);

            foreach (SettingsProperty currentProperty in Properties.Settings.Default.Properties)
            {
                DataRow row = dtConfig.NewRow();

                row["columnTitle"] = currentProperty.Name;
                row["columnValue"] = Properties.Settings.Default[currentProperty.Name];
                dtConfig.Rows.Add(row);
            }

            dtConfig = SetTableFilter(dtConfig, "columnTitle", "");
            dataGridView1.DataSource = dtConfig;
        }

        public static DataTable SetTableFilter(DataTable _objTable, string _Sort, String _Filter)
        {
            DataTable objTable = new DataTable();

            try
            {
                DataView dvFilter = _objTable.DefaultView;

                if (_objTable.Rows.Count > 0)
                {
                    dvFilter.Sort = _Sort;
                    dvFilter.RowFilter = _Filter;
                }

                if (dvFilter.Count > 0)
                {
                    objTable = dvFilter.ToTable();
                }
            }
            catch
            { }

            return objTable;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dtConfig.AcceptChanges();

                string path = Application.ExecutablePath + ".config";
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                if (File.Exists(path))
                {
                    doc.Load(path);

                    foreach (DataRow row in dtConfig.Rows)
                    {
                        string configString = @"configuration/applicationSettings/" + DoService.pubService + ".Properties.Settings/setting[@name='" + row["ColumnTitle"].ToString() + "']/value";
                        System.Xml.XmlNode configNode = doc.SelectSingleNode(configString);

                        if (configNode != null)
                        {
                            configNode.InnerText = row["columnValue"].ToString();
                            Console.WriteLine(configNode.InnerText);
                        }
                    }
                    doc.Save(path);
                    Properties.Settings.Default.Reload();
                }

                this.DialogResult = DialogResult.OK;
            }
            catch
            {
                this.DialogResult = DialogResult.Cancel;
            }
        }
    }
}
