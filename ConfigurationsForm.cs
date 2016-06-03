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
    public partial class frmConfigurations : Form
    {
        public frmConfigurations()
        {
            InitializeComponent();
            cmbCompany.SelectedIndexChanged -= OnCompanyChanged;
            txtOutputPath.Text = Properties.Settings.Default.OutputPath;
            txtFormat.Text = Properties.Settings.Default.FolderNameFormat;
            txtContest.Text = Properties.Settings.Default.ContestName;
            txtCustNo.Text = Properties.Settings.Default.CustCareNo;
            cmbCompany.DataSource = Enum.GetNames(typeof(Company));
            cmbCompany.SelectedItem = Properties.Settings.Default.Company;
            cmbCompany.SelectedIndexChanged += OnCompanyChanged;

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog browser = new FolderBrowserDialog())
            {
                if (browser.ShowDialog() == DialogResult.OK)
                    txtOutputPath.Text = browser.SelectedPath;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string errorMessage = string.Empty;
            if (IsValidData(out errorMessage))
            {
                try
                {
                    Properties.Settings.Default.OutputPath = txtOutputPath.Text;
                    Properties.Settings.Default.FolderNameFormat = txtFormat.Text.Trim();
                    Properties.Settings.Default.ContestName = txtContest.Text.Trim();
                    Properties.Settings.Default.CustCareNo = txtCustNo.Text.Trim();
                    Properties.Settings.Default.Company = cmbCompany.SelectedItem.ToString();

                    Properties.Settings.Default.Save();
                    Constants.LoadSettings();
                    MessageBox.Show("Setting saved successfully", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to save settings \n" + ex.Message, 
                        this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show(errorMessage, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsValidData(out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {

                DateTime tempDt = DateTime.Now;
                if (string.IsNullOrEmpty(txtOutputPath.Text))
                    errorMessage = "Output path cannot be empty";
                if (string.IsNullOrEmpty(txtFormat.Text.Trim()))
                    errorMessage = "Folder name format cannot be empty";
                if (!DateTime.TryParse(tempDt.ToString(txtFormat.Text.Trim()), out tempDt))
                    errorMessage = "Folder name format is not valid";
                if (string.IsNullOrEmpty(txtCustNo.Text.Trim()))
                    errorMessage = "Custmor care number cannot be empty";
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }

            return errorMessage.Length == 0;
        }

        private void OnCompanyChanged(object sender, EventArgs e)
        {
            MessageBox.Show("You are changing the company. Make sure that all settings like customer care number, logo etc are correct.", 
                this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
