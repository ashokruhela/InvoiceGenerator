using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceGenerator
{
    public partial class frmInvoiceGen : Form
    {
        string filePath = string.Empty;
        DataTable dtExcelData = new DataTable();
        ExcelDataProvider excelDataProvider = null;

        public frmInvoiceGen()
        {
            InitializeComponent();
            AddThemes();
            cmbTheme.SelectedItem = Properties.Settings.Default.BackColor;
            tspLabel.Text = "Shows the progress";
            excelDataProvider = new ExcelDataProvider();
            excelDataProvider.OpenExcelFile = excelDataProvider.OpenExcelFile;
            excelDataProvider.UpdateProgress += OnUpdateProgress;
            SetTitle();
        }

        void OnUpdateProgress(object sender, string currentValue)
        {
            Action updatelabel = new Action(() => { tspLabel.Text = currentValue; });
            UpdateControlValue(tspLabel.GetCurrentParent(), updatelabel);
        }

        private void UpdateControlValue(Control control, Action action)
        {
            control.Invoke(action);
        }

        private async void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel Files (*.xls, *.xls)|*.xls;*.xlsx";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    tspLabel.Text = "Loading excel file....";
                    txtFilePath.Text = filePath = dialog.FileName;
                    await LoadExcelFileAsync(filePath);
                    tspLabel.Text = "Excel file loaded";
                }
            }
            catch(Exception ex)
            {
                tspLabel.Text = "Shows the progress";
                MessageBox.Show(string.Format("Error in generating invoice\n.{0}", ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private async void btnRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                if (filePath.Length <= 0)
                    MessageBox.Show(this, "Select excel file to generate invoice", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    tspLabel.Text = "Loading excel file....";
                    txtFilePath.Text = filePath;
                    await LoadExcelFileAsync(filePath);
                    tspLabel.Text = "Data loaded";
                }
            }
            catch (Exception ex)
            {
                tspLabel.Text = "Show the progress";
                MessageBox.Show(this,string.Format("Error in generating invoice\n.{0}", ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            if (dtExcelData.Rows.Count == 0)
            {
                MessageBox.Show(this,"Select excel file to generate invoice", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    btnGenerate.Enabled = false;
                    btnBrowse.Enabled = false;
                    btnRefresh.Enabled = false;
                    tspLabel.Text = "Generating invoice....";
                    bool isInvoceGenerated = await excelDataProvider.GenerateInvoice(dtExcelData);
                    if (!isInvoceGenerated)
                    {
                        MessageBox.Show(this,"Error in generating invoice.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tspLabel.Text = "Error in generating invoice";
                    }
                    else
                    {
                        tspLabel.Text = "Invoice generated successfully";
                        MessageBox.Show(this,"Invoice generated successfully", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                finally
                {
                    btnGenerate.Enabled = true;
                    btnBrowse.Enabled = true;
                    btnRefresh.Enabled = true;
                }
            }
        }

        private async Task LoadExcelFileAsync(string filePath)
        {
            if (filePath.Trim().Length > 0)
            {
                dtExcelData.Rows.Clear();
                dtExcelData = await excelDataProvider.GetExcelDataAsync(filePath);
                dgvData.DataSource = dtExcelData;

                foreach (DataColumn dtColumn in dtExcelData.Columns)
                {
                    dgvData.Columns[dtColumn.ColumnName].HeaderText = dtColumn.Caption;
                    //if (dtColumn.Caption == Constants.Skip)
                    //    dgvData.Columns[dtColumn.ColumnName].Visible = false;
                }
            }
            
        }

        private void chkOpenFile_CheckedChanged(object sender, EventArgs e)
        {
            excelDataProvider.OpenExcelFile = chkOpenFile.Checked;
        }

        private void AddThemes()
        {
            foreach (Color color in new ColorConverter().GetStandardValues())
            {
                if (color.Name.ToUpper() != "TRANSPARENT")
                    cmbTheme.Items.Add(color.Name);
            }
        }

        private void cmbTheme_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyTheme(cmbTheme.SelectedItem.ToString());
        }

        private void ApplyTheme(string color)
        {
            this.BackColor = Color.FromName(color);
            statusStrip1.BackColor = this.BackColor;
            Properties.Settings.Default.BackColor = this.BackColor.Name;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
            base.OnClosing(e);
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            frmConfigurations configForm = new frmConfigurations();
            configForm.BackColor = this.BackColor;
            configForm.StartPosition = FormStartPosition.CenterScreen;
            configForm.ShowDialog();
            SetTitle();
        }
        
        public void SetTitle()
        {
            this.Text = string.Format("Invoice generator - {0}", Constants.CompanyName.ToString());
        }
    }
}
