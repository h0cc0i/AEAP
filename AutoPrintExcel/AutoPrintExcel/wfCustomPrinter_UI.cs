using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using Spire.Pdf;
using System.Globalization;
using System.Resources;

namespace AutoPrintExcel
{
    public partial class wfCustomPrinter_UI : Form
    {
        //2016/07/25 change default tray with form in VN
        const int DEFAULTTRAYINDEX = 2;
        const int DEFAULTDUPLEXINDEX = 1;

        //2016/07/15_HonC Define DefaultPrinterSetting Global Variable;
        private PrintDialog _DefaultPrinterSettings = new PrintDialog();


        //2016/07/27 _HonC
        CultureInfo culture;

        public wfCustomPrinter_UI()
        {
            InitializeComponent();
        }


        /// <summary>
        /// Load Default Printer Settings and Show in form
        /// </summary>
        private void LoadPrinterSettings()
        {
            //Define printer settings
            PrinterSettings _settings = new PrinterSettings();
            lblPrinterName.Text = _settings.PrinterName;

            //Get Printer settings tray
            cmbTray.DataSource = LibPrintExcel.GetTrayofPrinter();
            cmbTray.DisplayMember = "TrayName";
            cmbTray.ValueMember = "TrayIndex";

            //Get Printer duplex
            cmbDuplex.DataSource = LibPrintExcel.DefaultDuplexCombo();
            cmbDuplex.DisplayMember = "DuplexName";
            cmbDuplex.ValueMember = "DuplexID";

        }

        private void wfCustomPrinter_UI_Load(object sender, EventArgs e)
        {
            try
            {
                LoadPrinterSettings();
                //2016/07/27 _HonC load default Language
                LoadLaguage(LibStub._DefaultLanguage);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void Quít_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Set default setting for Printer
        private void btnUseDefaultSettings_Click(object sender, EventArgs e)
        {
            SaveDefaultSettings();
            LibPrintExcel._USECUSTOMSETTINGS = false;
        }


        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            //2016/07/21 _HonC set flag use Custom Settings
            LibPrintExcel._USECUSTOMSETTINGS = true;

            Int32 i = 0;
            try
            {
                //Set default Printer settings
                LibPrintExcel._DEFAULTSETTINGS = LibPrintExcel.SetPrinterSettings(cmbTray.Text.ToString(), LibPrintExcel.ReturnDuplex(cmbDuplex.SelectedIndex));

                //2016/07/19 _HonC set Default 
                LibPrintExcel._FROMPAGE = Int32.TryParse(txtfromPage.Text, out i) ? Int32.Parse(txtfromPage.Text) : 1;
                LibPrintExcel._TOPAGE = Int32.TryParse(txtToPage.Text, out i) ? Int32.Parse(txtToPage.Text) : 1;

                LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.FromPage = LibPrintExcel._FROMPAGE;
                LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.ToPage = LibPrintExcel._TOPAGE;

                this.Close();
            }
            catch (Exception)
            {

                throw;
            }


        }

        /// <summary>
        /// 2016/07/20 _HonC
        /// Save Default Settings for Print
        /// </summary>
        public void SaveDefaultSettings()
        {
            cmbTray.SelectedIndex = DEFAULTTRAYINDEX;
            cmbDuplex.SelectedIndex = DEFAULTDUPLEXINDEX;
        }


        /// <summary>
        /// 2016/07/27 _HonC
        /// 
        /// </summary>
        public void LoadLaguage(string cultureName)
        {
            culture = CultureInfo.CreateSpecificCulture(cultureName);
            ResourceManager rm = new ResourceManager("AutoPrintExcel.Lang.MyResource", typeof(Form1).Assembly);

            lblFrom.Text = rm.GetString("From", culture);
            lblPrintDuplex.Text = rm.GetString("PrintDubplex", culture);
            //lblPrinterName.Text = rm.GetString("Name", culture);
            lblPrinterSettings.Text = rm.GetString("PrintSetting", culture);
            lblPrintNameLable.Text = rm.GetString("PrinterName", culture);
            lblPrintPages.Text = rm.GetString("PrintPages", culture);
            lblTo.Text = rm.GetString("To", culture);
            lblTray.Text = rm.GetString("Tray", culture);

            btnSaveSettings.Text = rm.GetString("SaveSettings", culture);
            btnUseDefaultSettings.Text = rm.GetString("DefaultSettings", culture);
            btnQuit.Text = rm.GetString("Quit", culture);
        }
    }
}
