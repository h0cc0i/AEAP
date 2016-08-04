using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.Globalization;
using System.Resources;
using AutoPrintExcel.Button_Class;
using Microsoft.Win32;

namespace AutoPrintExcel
{
    public partial class Form1 : Form
    {
        CultureInfo culture;
        string[] _listFileName = new string[99];
        string[] _listFilePath = new string[99];
        System.Data.DataTable _dtbTreeView = new System.Data.DataTable();
        Microsoft.Office.Interop.Excel.Application xlexcel;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        int _flagcount = 0;
        Image bmp;
        //define DataTable for List Hino
        System.Data.DataTable _dtbListHino = new System.Data.DataTable();

        //2016/07/11_Honc  define list tray Name of default printer
        List<string> _List = new List<string>();

        //2016/07/11_HonC define Stream reader for Print
        private StreamReader _StrReader;
        private System.Drawing.Font printFont;

        //2016/07/20 _HonC define DataTable for Search Engine
        System.Data.DataTable _DataSourcee;

        #region "create flag for button check"
        bool _flagbtnBKM = false;
        bool _flagbtnRRC = false;
        bool _flagbtnHino = false;
        #endregion

        //declare a variable to hold the CurrentCulture
        System.Globalization.CultureInfo oldCI;
        //get the old CurrenCulture and set the new, en-US
        void SetNewCurrentCulture()
        {
            oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }
        //reset Current Culture back to the originale
        void ResetCurrentCulture()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
        }


        System.Data.DataTable _dtbExcel = new System.Data.DataTable();
        public Form1()
        {
            ////create cursor
            //Cursor _mCursor = new System.Windows.Forms.Cursor("\\Mouse Source\\SmoothHourglass.ani");
            InitializeComponent();
            culture = CultureInfo.CurrentCulture;
            btnVNLanguage.Checked = true;
            btnEnglishLanguage.Checked = false;
            SetNewCurrentCulture();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        // method get sheet name of excel workbook
        private void GetExcelSheetNames(string excelfile, string excelparent)
        {


            OleDbConnection ObjConn = null;
            System.Data.DataTable dtb = null;
            string connString = "";
            //check xls or xlsx
            if (excelfile.Substring(excelfile.Length - 4) == "xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
              "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 Xml;HDR=YES;Readonly=False;\"";
            }
            else
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                      "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            }


            try
            {
                //create connection string
                ObjConn = new OleDbConnection(connString);
                ObjConn.Open();
                // get data table cotaining the schema 
                dtb = ObjConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                String[] _excelSheet = new string[dtb.Rows.Count];
                int i = 0;

                // loop get sheet name
                foreach (DataRow _row in dtb.Rows)
                {
                    _excelSheet[i] = _row["TABLE_NAME"].ToString();
                    i++;
                }

                int _excelSheetcount = _excelSheet.Length;
                TreeNode[] _arraay = new TreeNode[_excelSheetcount];
                //loop add sheet name to treeview
                for (int j = 0; j < _excelSheetcount; j++)
                {
                    TreeNode _treeson = new TreeNode(_excelSheet[j].ToString().Replace("'", ""));
                    _treeson.Tag = excelfile;
                    _arraay[j] = _treeson;
                }

                //create tree parent
                if (_arraay != null)
                {
                    TreeNode _treeparent = new TreeNode(excelparent, _arraay);
                    _treeparent.Tag = excelfile;
                    //add to treeview
                    treeView1.Nodes.Add(_treeparent);
                }
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.ToString());
                lbnMessage.Text = "Get Excel and Sheet failed, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
            finally // close and dispose connection
            {
                ObjConn.Close();
                ObjConn.Dispose();
            }
        }


        //method add name of pdf to treeview
        private void AddpdfintoTreeView(string _pdffile, string _pdfname)
        {
            TreeNode _treeparent = new TreeNode(_pdfname);
            _treeparent.Tag = _pdffile;
            //add to treeview
            treeView1.Nodes.Add(_treeparent);
        }

        // method read file excel to datagridview
        private System.Data.DataTable GetDataTable(string excelfile, string excelparent)
        {
            System.Data.DataTable dtb = new System.Data.DataTable();

            OleDbConnection ObjConn = null;
            System.Data.DataTable dtbExcel = new System.Data.DataTable();
            string connString = "";
            #region "check xls or xlsx"
            //check xls or xlsx
            if (excelfile.Substring(excelfile.Length - 4) == "xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
              "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";
            }
            else
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                      "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            }
            #endregion

            try
            {
                ObjConn = new OleDbConnection(connString);
                ObjConn.Open();

                // get data table cotaining the schema 
                dtb = ObjConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string SheetName = dtb.Rows[0]["TABLE_NAME"].ToString();
                string CommandText = "SELECT * From [" + SheetName + "]";

                OleDbCommand cmdExcel = new OleDbCommand(CommandText, ObjConn);
                OleDbDataAdapter oda = new OleDbDataAdapter(CommandText, ObjConn);

                oda.SelectCommand = cmdExcel;
                oda.Fill(dtbExcel);
                return dtbExcel;

            }
            catch (Exception)
            {
                lbnMessage.Text = "Get Excel and Sheet failed, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
            finally // close and dispose connection
            {
                ObjConn.Close();
                ObjConn.Dispose();
            }
            return dtb;
        }

        // method get value of cell
        private void SetTextforCell(string _pathexcel, string _sheetName)
        {

            object misValue = System.Reflection.Missing.Value;


            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel);

            // go to sheet 
            Microsoft.Office.Interop.Excel.Sheets _excelsheets = wb.Worksheets;

            //get contains excel
            foreach (Worksheet _ws in wb.Worksheets)
            {
                _ws.Name = _ws.Name.ToString().Trim();
                _sheetName = _sheetName.Trim();
                if ((_ws.Name.ToString() == _sheetName) || (_ws.Name.Contains(_sheetName)))
                {
                    Microsoft.Office.Interop.Excel.Worksheet _excelworksheet = _excelsheets.get_Item(_ws.Name.ToString());
                    _excel.DisplayAlerts = false; // disable alert message box 

                    // auto set text with textbox txtCell
                    _excelworksheet.Range[txtCell.Text].set_Value(misValue, txtEdit.Text);

                    #region "code Save file Excel"

                    // save file excel with get type of file
                    //if (_pathexcel.Substring(_pathexcel.Length - 4) == "xlsx")
                    //{
                    //    wb.SaveAs(_pathexcel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                    //        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                    //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
                    //        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    //}
                    //else
                    //{
                    //    wb.SaveAs(_pathexcel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing,
                    //       Type.Missing, false, misValue, XlSaveAsAccessMode.xlExclusive,
                    //       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //}
                    #endregion

                    wb.Save();
                    #region  "Change printer size"
                    // Get the current printer

                    string Defprinter = null;
                    Defprinter = xlexcel.ActivePrinter;
                    PrinterSettings _pr = new PrinterSettings();
                    var _with1 = xlWorkSheet.PageSetup;
                    // Leter papersize
                    _with1.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperLetter;
                    #endregion

                    xlexcel.ActivePrinter = _pr.PrinterName;

                    _excelworksheet.PrintOut(Type.Missing, 1, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                    wb.Close();
                    _workbook.Close();
                    _excel.DisplayAlerts = true;
                    _excel.Quit();

                    // release all
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
                    _excel = null;

                    break;
                }
            }
            //kill all process !!!
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.ProcessName == "EXCEL")
                    process.Kill();
            }

        }

        //method just print with sheetname and path excel --| 2015/12/14
        private void PrintWithSheetName(string _pathexcel, string _sheetName)
        {
            _sheetName = _sheetName.Trim();
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel);

            // go to sheet 
            Microsoft.Office.Interop.Excel.Sheets _excelsheets = wb.Worksheets;

            //get contains excel
            foreach (Worksheet _ws in wb.Worksheets)
            {
                if (_ws.Name.Contains(_sheetName))
                {
                    Microsoft.Office.Interop.Excel.Worksheet _excelworksheet = _excelsheets.get_Item(_ws.Name.ToString());
                    _excel.DisplayAlerts = false; // disable alert message box 

                    // save file excel with get type of file
                    //if (_pathexcel.Substring(_pathexcel.Length - 4) == "xlsx")
                    //{
                    //    wb.SaveAs(_pathexcel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                    //        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                    //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
                    //        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    //}
                    //else
                    //{
                    //    wb.SaveAs(_pathexcel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing,
                    //       Type.Missing, false, misValue, XlSaveAsAccessMode.xlExclusive,
                    //       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //}


                    //wb.Save();

                    // print excel with first page
                    // _excelworksheet.PageSetup.PrintArea = "A1:G38";

                    _excelworksheet.PrintOut(Type.Missing, 1, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                    wb.Close();
                    _workbook.Close();
                    _excel.DisplayAlerts = true;
                    _excel.Quit();

                    // release all
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
                    _excel = null;

                    // kill all process name Excel
                    Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
                    if (!excelProcess.CloseMainWindow())
                    {
                        excelProcess.Kill();
                    }
                    break;

                }
            }

        }

        //method just print, auto print active sheet with path excel
        static void JustPrint(string _pathexcel)
        {
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel);   //fix

            // go to sheet 
            Microsoft.Office.Interop.Excel.Worksheet _excelsheets = wb.Worksheets[1];
            //Microsoft.Office.Interop.Excel.Worksheet _excelworksheet = _excelsheets.get_Item(wb.ActiveSheet);

            _excel.DisplayAlerts = false; // disable alert message box 


            //PrinterSettings _setting = new PrinterSettings();
            //foreach (System.Drawing.Printing.PrinterSettings _printer in PrinterSettings.InstalledPrinters)

            var printer = System.Drawing.Printing.PrinterSettings.InstalledPrinters;


            //_excelsheets.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            _excelsheets.PageSetup.PaperSize = XlPaperSize.xlPaperA4;


            _excelsheets.PrintOut(Type.Missing, 1, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            wb.Close();
            _workbook.Close();
            _excel.DisplayAlerts = true;
            _excel.Quit();

            // release all
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
            _excel = null;



            // kill all process name Excel
            Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
            if (!excelProcess.CloseMainWindow())
            {
                excelProcess.Kill();
            }

        }

        //2016/07/11_HonC function auto print excel file with tray name of printer
        void JustPrint(string _pathexcel, string _Tray)
        {
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook _wb = _workbook.Open(_pathexcel);
            _wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF);
            // go to sheet 
            Microsoft.Office.Interop.Excel.Worksheet _excelsheets = _wb.Worksheets[1];
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            _excelsheets.Protect(Contents: false);

            //string startRange = "A1";
            //Microsoft.Office.Interop.Excel.Range endRange = _excelsheets.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //Range range = _excelsheets.get_Range(startRange, endRange);

            Range r = _excelsheets.Range["A1:G44"];
            r.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

            Bitmap image = new Bitmap(Clipboard.GetImage());


            PrintExcelWithBitMap(image, _Tray);


            _excel.DisplayAlerts = false; // disable alert message box 


            //_excelsheets.PrintOut(Type.Missing, 1, Type.Missing, Type.Missing,
            //Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            _wb.Close();
            _workbook.Close();
            _excel.DisplayAlerts = true;
            _excel.Quit();

            // release all
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_wb);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
            _excel = null;



            // kill all process name Excel
            Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
            if (!excelProcess.CloseMainWindow())
            {
                excelProcess.Kill();
            }

        }

        //method set language       --| 2015/12/15
        private void SetLanguage(string cultureName)
        {
            culture = CultureInfo.CreateSpecificCulture(cultureName);
            ResourceManager rm = new ResourceManager("AutoPrintExcel.Lang.MyResource", typeof(Form1).Assembly);
            btnFile.Text = rm.GetString("File", culture);
            btnImportDataSource.Text = rm.GetString("ImportDataSource", culture);
            btnImportListExe.Text = rm.GetString("ImportListExecute", culture);
            btnExit.Text = rm.GetString("Exit", culture);
            btnFuntion.Text = rm.GetString("Function", culture);
            btnclearTreeView.Text = rm.GetString("ClearTreeView", culture);
            btnClearGridView.Text = rm.GetString("ClearGridview", culture);
            btnFixNameKensa.Text = rm.GetString("FixNameKensa", culture);
            btnJustGetName.Text = rm.GetString("JustGetName", culture);
            btnEditAndPrint.Text = rm.GetString("EditandPrint", culture);
            btnGetPathCheckSheet.Text = rm.GetString("GetPathCheckSheet", culture);
            btnGetPathKensa.Text = rm.GetString("GetPathKensa", culture);

            //2016/08/01 _HonC
            btnCustomPrintMenu.Text = rm.GetString("PrintCustom", culture);

            btnJustPrint.Text = rm.GetString("JustPrint", culture);
            btnPrintKensa.Text = rm.GetString("PrintKensa", culture);
            tpListDataSource.Text = rm.GetString("ListDataSource", culture);
            tpListPrint.Text = rm.GetString("ListPrint", culture);
            tpListNotFound.Text = rm.GetString("ListNotFound", culture);
            tpListHino.Text = rm.GetString("ListSourceHino", culture);
            ckbEditBox.Text = rm.GetString("EditNameofCheckSheet", culture);

            btnHelp.Text = rm.GetString("Help", culture);
            btnLanguage.Text = rm.GetString("Language", culture);
            btnVNLanguage.Text = rm.GetString("VietNamese", culture);
            btnEnglishLanguage.Text = rm.GetString("English", culture);
            btnAbout.Text = rm.GetString("About", culture);
            lbnText.Text = rm.GetString("Text", culture);
            lbnCell.Text = rm.GetString("Cell", culture);


            //2016/07/27 _HonC
            // Add define language 
            btnPdfDataSource.Text = rm.GetString("ImportPDFSource", culture);
            btnClearListHino.Text = rm.GetString("ClearListHino", culture);
            btnFixNameCheckSheet.Text = rm.GetString("FixNameCheckSheet", culture);
            btnGetPathDrawing.Text = rm.GetString("GetPathDrawing", culture);
            btnPrintDrawing.Text = rm.GetString("PrintDrawing", culture);
            btnImportListDSNew.Text = rm.GetString("ImportListDataSource", culture);
            btnimportListExcuteNEW.Text = rm.GetString("ImportListExecute", culture);
            btngetPathNEW.Text = rm.GetString("GetPath", culture);
            btnprintCustom.Text = rm.GetString("PrintCustom", culture);
            btnCustomPrinter.Text = rm.GetString("CustomPrinter", culture);
            btnExecutePDFFile.Text = rm.GetString("ExecutePDFFile", culture);
            btnImportListDefineHino.Text = rm.GetString("ImportListDefineHino", culture);


        }

        private void AutoPrint(string _pathexcel, string _sheetName)
        {
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel);   //fix

            // go to sheet 
            Microsoft.Office.Interop.Excel.Sheets _excelsheets = wb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet _excelworksheet = _excelsheets.get_Item(_sheetName);

            _excel.DisplayAlerts = false; // disable alert message box 

            // print excel with first page
            // _excelworksheet.PageSetup.PrintArea = "A1:G38";

            _excelworksheet.PrintOut(Type.Missing, 1, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //_excelworksheet.PrintOutEx(Type.Missing, 1, Type.Missing, Type.Missing,
            //Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            wb.Close();
            _workbook.Close();
            _excel.DisplayAlerts = true;
            _excel.Quit();

            // release all
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
            _excel = null;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            try
            {

                // Open File Dialog
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = true;

                if (_open.ShowDialog() == DialogResult.OK)
                {
                    // txtFilePath.Text = _open.SafeFileName;
                    _listFileName = _open.SafeFileNames;

                    //add list to tree view
                    int _countList = _listFileName.Count();
                    for (int i = 0; i < _countList; i++)
                    {
                        GetExcelSheetNames(_open.FileNames[i].ToString(), _open.SafeFileNames[i].ToString());
                    }
                }



            }
            catch (Exception)
            {
                lbnMessage.Text = "Open file excel failed, please try later.";
                lbnMessage.ForeColor = Color.Red;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //2016/07/27 _HonC Set default Language
            //LibStub._DefaultLanguage = "vi-VN";
            SetLanguage(LibStub._DefaultLanguage);

            lbnMessage.Text = "Auto Edit and Print Excel v1.4";
            lbnMessage.ForeColor = Color.Blue;
            lbnQuantity.Text = string.Empty;
            pnlEditName.Visible = false;
            ckbEditBox.Checked = false;

            //2016/07/15_HonC  Clear Path of DEFAUTLPDFTEMP
            LibPrintExcel._DEFAULTPDFTEMPPATH = string.Empty;

            SetDefaultButton(true, false, false);
            // kill all process name Excel
            //Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
            //if (!excelProcess.CloseMainWindow())
            //{
            //    excelProcess.Kill();
            //}

            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.ProcessName == "EXCEL")
                    process.Kill();
            }

        }

        private void SetDefaultButton(bool _flagBKM, bool _flagRRC, bool _flagHino)
        {
            //set btn BeckMan
            _flagbtnBKM = _flagBKM;
            if (_flagBKM == true)
            {
                btnBKM.ForeColor = Color.White;
                btnBKM.BackColor = Color.DarkOrange;
            }
            else
            {
                btnBKM.ForeColor = Color.Black;
                btnBKM.BackColor = Color.LightBlue;

            }

            //set button RRC
            _flagbtnRRC = _flagRRC;
            if (_flagRRC == true)
            {
                btnRRC.ForeColor = Color.White;
                btnRRC.BackColor = Color.DarkOrange;
            }
            else
            {
                btnRRC.ForeColor = Color.Black;
                btnRRC.BackColor = Color.LightBlue;
            }

            //set button Hino
            _flagbtnHino = _flagHino;
            if (_flagbtnHino == true)
            {
                btnHino.ForeColor = Color.White;
                btnHino.BackColor = Color.DarkOrange;

            }
            else
            {
                btnHino.ForeColor = Color.Black;
                btnHino.BackColor = Color.LightBlue;
            }
        }


        private void btnEditCell_Click(object sender, EventArgs e)
        {
            try
            {

                // Check checkbox is checked
                foreach (TreeNode _node in treeView1.Nodes)
                {
                    foreach (TreeNode _childNode in _node.Nodes)
                    {
                        if (_childNode.Checked == true)
                        {
                            string _strchidNode = _childNode.Text.Replace("'", "");
                            _strchidNode = _strchidNode.Replace("$", "");
                            //if (File.Exists(_childNode.Tag.ToString()))
                            //{
                            //    File.Delete(_childNode.Tag.ToString());
                            //}
                            SetTextforCell(_childNode.Tag.ToString(), _strchidNode);
                        }
                    }

                }
            }
            catch (Exception)
            {

                lbnMessage.Text = "Auto Edit Excel failed, please try later.";
                lbnMessage.ForeColor = Color.Red;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            lbnMessage.Text = "Auto Edit and Print Excel v1.4";
            lbnMessage.ForeColor = Color.Blue;
            lbnQuantity.Text = string.Empty;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {

                // Check checkbox is checked
                foreach (TreeNode _node in treeView1.Nodes)
                {
                    foreach (TreeNode _childNode in _node.Nodes)
                    {
                        if (_childNode.Checked == true)
                        {
                            string _strchidNode = _childNode.Text.Replace("'", "");
                            _strchidNode = _strchidNode.Replace("$", "");

                            AutoPrint(_childNode.Tag.ToString(), _strchidNode);
                        }
                    }

                }

                lbnMessage.Text = "Auto print multiple excel failed, please try later.";
                lbnMessage.ForeColor = Color.Red;
            }
            catch (Exception)
            {
                lbnMessage.Text = "Auto print multiple excel failed, please try later.";
                lbnMessage.ForeColor = Color.Red;
            }
        }

        private void mnStripImportDataSource_Click(object sender, EventArgs e)
        {

            if (_dtbTreeView.Columns.Count < 1)
            {
                _dtbTreeView.Columns.Add("Name");
                _dtbTreeView.Columns.Add("Path");
            }
            else
            {
                treeView1.Nodes.Clear();
            }

            try
            {
                // Open File Dialog
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = true;

                if (_open.ShowDialog() == DialogResult.OK)
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    // txtFilePath.Text = _open.SafeFileName;
                    _listFileName = _open.SafeFileNames;


                    //add list to tree view
                    int _countList = _listFileName.Count();
                    for (int i = 0; i < _countList; i++)
                    {
                        GetExcelSheetNames(_open.FileNames[i].ToString(), _open.SafeFileNames[i].ToString());
                        //add data to temp table

                        _dtbTreeView.Rows.Add(_open.SafeFileNames[i].ToString(), _open.FileNames[i].ToString());
                    }

                    // go to treeview tab
                    lbnQuantity.Text = _countList.ToString();
                    lbnMessage.Text = "Import data source successful";
                    lbnMessage.ForeColor = Color.Blue;

                }
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.ToString());
                lbnMessage.Text = "Import data source failed. ";
                lbnMessage.ForeColor = Color.Red;

            }
            Cursor.Current = Cursors.Default;
        }

        private void mnStripImportListToPrint_Click(object sender, EventArgs e)
        {
            if (_flagbtnHino == false || _dtbListHino.Rows.Count > 0)
            {
                _dtbExcel = new System.Data.DataTable();
                dtgListPrint.DataSource = new System.Data.DataTable();
                dtgError.DataSource = new System.Data.DataTable();
                // import file excel -> read excel to datagridview
                try
                {
                    // Open File Dialog
                    OpenFileDialog _open = new OpenFileDialog();
                    _open.Filter = "All Files (*.*)|*.*";
                    _open.FilterIndex = 1;
                    _open.Multiselect = false;

                    if (_open.ShowDialog() == DialogResult.OK)
                    {
                        _listFileName = _open.SafeFileNames;
                        //System.Data.DataTable dtb = new System.Data.DataTable();
                        //read excel to datagriview
                        _dtbExcel = GetDataTable(_open.FileName.ToString(), _open.SafeFileName.ToString());
                        if ((_dtbExcel != null) && _dtbExcel.Rows.Count > 0)
                            dtgListPrint.DataSource = _dtbExcel;

                        dtgListPrint.Columns["Path"].Width = 500;

                        for (int i = 0; i < _dtbExcel.Rows.Count; i++)
                        {
                            if ((_dtbExcel.Rows[i]["Name"]) == null || _dtbExcel.Rows[i]["Name"].ToString() == "")
                                _dtbExcel.Rows[i].Delete();
                        }
                        _dtbExcel.AcceptChanges();
                    }
                    // go to treeview tab
                    //tabControl1.SelectedTab = tabPage2;
                    lbnMessage.Text = "Import list name of check sheet successful";
                    lbnMessage.ForeColor = Color.Blue;
                    lbnQuantity.Text = _dtbExcel.Rows.Count.ToString();
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Import list name of check sheet failed";
                    lbnMessage.ForeColor = Color.Red;
                    lbnMessage.Text = "";
                }
            }
            else
            {
                lbnMessage.Text = "Please insert List Define of Hino first and try it later.";
                lbnQuantity.Text = "";
            }
        }

        private void btnGetPathExcel_Click(object sender, EventArgs e)
        {
            //define flag break out of foreach loop
            bool _keeploop = true;
            //count row in dtb list execute
            int _Countdtb = _dtbExcel.Rows.Count;

            #region "Old Code- delete because when import list ~~ delete row null or row emty int dtbExcel"
            //for (int i = 0; i < _dtbExcel.Rows.Count; i++)
            //{
            //    if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Name"].ToString()) || _dtbExcel.Rows[i]["Name"].ToString() == "")
            //    {
            //        _dtbExcel.Rows[i].Delete();
            //    }
            //}
            //_dtbExcel.AcceptChanges();
            #endregion


            _Countdtb = dtgListPrint.Rows.Count;
            int _flagbug = 0;


            // Check Name by Treeview --> Get Path
            if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
            {
                try
                {
                    if (!_flagbtnHino)
                    {
                        for (int i = 0; i < _Countdtb; i++)
                        {
                            if (_keeploop)
                            {
                                foreach (TreeNode _node in treeView1.Nodes)
                                {
                                    if (_keeploop)
                                    {
                                        foreach (TreeNode _childNode in _node.Nodes)
                                        {
                                            string _strnode = _childNode.Text.Replace("'", "").Replace("$", "").Replace("Laze", "").Replace("Print_Area", "").Trim();
                                            string _RowExecute = _dtbExcel.Rows[i]["Name", DataRowVersion.Original].ToString().Trim();
                                            if (_flagbtnBKM)
                                            {
                                                // Thuat toan for BeckMan
                                                if ((_strnode == _RowExecute) || ((_strnode.Length > 6) && !(String.IsNullOrEmpty(_RowExecute)) && (_strnode.Remove(6) == _RowExecute.Trim())))
                                                {
                                                    _dtbExcel.Rows[i]["Path"] = _childNode.Tag.ToString();
                                                    _flagbug++;
                                                    _keeploop = false;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                // Thuat toan for RRC
                                                if ((_strnode.Length > 6) && !(String.IsNullOrEmpty(_RowExecute)) && (_strnode == _RowExecute.Trim()))
                                                {
                                                    _dtbExcel.Rows[i]["Path"] = _childNode.Tag.ToString();
                                                    _flagbug++;
                                                }
                                            }


                                        }
                                    }
                                }
                                _keeploop = true;
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            //Execute list Hino
                            AutoCompareData(_dtbListHino, _dtbExcel);

                            dtgListPrint.DataSource = _dtbExcel;
                            dtgListPrint.Columns["ID"].DisplayIndex = 0;
                            //done create and search virtual list.

                            //search Path with ID
                            for (int i = 0; i < _Countdtb; i++)
                            {
                                foreach (TreeNode _node in treeView1.Nodes)
                                {
                                    foreach (TreeNode _childNode in _node.Nodes)
                                    {
                                        string _strnode = _childNode.Text.Replace("'", "").Replace("Laze", "").Replace("Laser", "").Replace("$", "").Replace(" ", "").Trim();
                                        string _RowExecute = _dtbExcel.Rows[i]["ID"].ToString().Trim();
                                        if (!(String.IsNullOrEmpty(_RowExecute)) && (_strnode == _RowExecute.Trim()))
                                        {
                                            _dtbExcel.Rows[i]["Path"] = _childNode.Tag.ToString();
                                            _flagbug++;
                                            break;
                                        }
                                    }
                                }
                            }

                        }
                        catch (Exception)
                        {
                            lbnText.Text = "Error. Please try later.";
                        }

                    }

                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;


                    // Check list not founds
                    System.Data.DataTable _dtbError = new System.Data.DataTable();
                    _dtbError.Columns.Add("Name");
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        if (string.IsNullOrEmpty(Convert.ToString(_dtbExcel.Rows[i]["Path"])))
                        {
                            // path is null --> delete this row
                            _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
                            _dtbExcel.Rows[i].Delete();
                        }
                    }
                    _dtbExcel.AcceptChanges();
                    if (_dtbError.Rows.Count > 0)
                    {
                        dtgError.DataSource = _dtbError;
                    }
                    lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
                    _dtbError = new System.Data.DataTable();
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed";
                    lbnMessage.ForeColor = Color.Red;
                }

            }
            else
            {
                lbnMessage.Text = "Please insert data source or list execute before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }
        //　お客様はＲＲＣの図面をさがしたいんです。

        #region "Old search engine"
        //bool _flag = false;

        //// Check Name by Treeview --> Get Path
        //if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
        //{
        //    try
        //    {
        //        for (int i = 0; i < _Countdtb; i++)
        //        {
        //            foreach (TreeNode _node in treeView1.Nodes)
        //            {
        //                foreach (TreeNode _childNode in _node.Nodes)
        //                {
        //                    string _str = _childNode.Text.Replace("'", "").Replace("$", "").Trim();
        //                    string _strdbExcelrow = _dtbExcel.Rows[i]["Name"].ToString().Trim();

        //                    if ((_strdbExcelrow == _str) || ((_strdbExcelrow + " Laser") == _str) || ((_strdbExcelrow + " Laze") == _str))
        //                    {
        //                        _dtbExcel.Rows[i]["Path"] = _childNode.Tag.ToString();
        //                        _flag = true;
        //                        break;
        //                    }
        //                }
        //                if (_flag) break;

        //            }
        //            _flag = false;
        //        }
        //        lbnMessage.Text = "Get path excel successful";
        //        lbnMessage.ForeColor = Color.Blue;


        //        // Check list not founds
        //        System.Data.DataTable _dtbError = new System.Data.DataTable();
        //        _dtbError.Columns.Add("Name");
        //        for (int i = 0; i < _Countdtb; i++)
        //        {
        //            if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
        //            {
        //                _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
        //            }
        //        }
        //        if (_dtbError.Rows.Count > 0)
        //        {
        //            dtgError.DataSource = _dtbError;
        //        }
        //        lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
        //        _dtbError = new System.Data.DataTable();
        //    }
        //    catch (Exception ex)
        //    {
        //        lbnMessage.Text = "Get path excel failed";
        //        lbnMessage.ForeColor = Color.Red;
        //    }
        //}
        //else
        //{
        //    lbnMessage.Text = "Please insert data source before.";
        //    lbnMessage.ForeColor = Color.Blue;
        //}
        #endregion

        private void mnStripbtnEditAndPrint_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            int _countExcel = _dtbExcel.Rows.Count;
            lbnQuantity.Text = string.Empty;
            if (!string.IsNullOrEmpty(txtEdit.Text))
            {
                try
                {
                    for (int i = 0; i < _countExcel; i++)
                    {
                        if (_dtbExcel.Rows[i]["Path"].ToString() != "")
                        {
                            if (!_flagbtnHino)
                            {
                                SetTextforCell(_dtbExcel.Rows[i]["Path"].ToString(), _dtbExcel.Rows[i]["Name"].ToString());
                            }
                            else
                                SetTextforCell(_dtbExcel.Rows[i]["Path"].ToString(), _dtbExcel.Rows[i]["ID"].ToString());

                            lbnMessage.Text = _dtbExcel.Rows[i]["Name"].ToString() + " is pending . . .";
                            lbnQuantity.Text = (i + 1).ToString() + " / " + _countExcel.ToString();
                            lbnMessage.ForeColor = Color.Blue;
                        }
                    }
                    lbnMessage.Text = "Auto edit and print successful";
                    lbnMessage.ForeColor = Color.Blue;
                    _dtbExcel = new System.Data.DataTable();
                }
                catch (Exception)
                {
                    _dtbExcel = new System.Data.DataTable();
                    lbnMessage.Text = "Auto edit and print fail, please try later";
                    lbnMessage.ForeColor = Color.Red;
                }
                this.BringToFront();
            }
            else
            {
                lbnMessage.Text = "Cell Edit is not null, please check again.";
                ckbEditBox.Checked = true;
                txtEdit.Focus();
            }

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow;
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int _countExcel = _dtbExcel.Rows.Count;
            lbnQuantity.Text = "";
            try
            {
                for (int i = 0; i < _countExcel; i++)
                {
                    if (!String.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                    {
                        // AutoPrint(_dtbExcel.Rows[i]["Path"].ToString(), _dtbExcel.Rows[i]["Name"].ToString());
                        lbnMessage.Text = _dtbExcel.Rows[i]["Name"].ToString() + " is printing . . .";
                        lbnQuantity.Text = (i + 1).ToString() + " / " + _countExcel.ToString();
                        lbnMessage.ForeColor = Color.Blue;
                        JustPrint(_dtbExcel.Rows[i]["Path"].ToString());
                    }
                }
                lbnMessage.Text = "Auto Print list kensa successful.";
                lbnMessage.ForeColor = Color.Blue;
                _dtbExcel = new System.Data.DataTable();
            }
            catch (Exception)
            {
                _dtbExcel = new System.Data.DataTable();
                lbnMessage.Text = " Auto print failed, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
            this.BringToFront();
        }

        private void clearGridView(object sender, EventArgs e)
        {
            dtgListPrint.DataSource = new System.Data.DataTable();
            dtgError.DataSource = new System.Data.DataTable();
            lbnMessage.Text = "Auto Edit and Print Excel v1.4";
            lbnMessage.ForeColor = Color.Blue;
            lbnQuantity.Text = string.Empty;
        }

        //Read and auto save as excel
        private void EditKensa(object sender, EventArgs e)
        {
            // Check Name by Treeview --> Get Path
            if (treeView1.Nodes.Count > 0 && !string.IsNullOrEmpty(txtCell.Text))
            {
                try
                {
                    foreach (TreeNode _node in treeView1.Nodes)
                    {
                        _flagcount++;

                        //GetNameKensa(_childNode.Tag.ToString(), _childNode.Text);
                        AutoSaveEachSheetKenSa(_node.Tag.ToString());

                        // kill all process name Excel
                        Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
                        if (!excelProcess.CloseMainWindow())
                        {
                            excelProcess.Kill();
                        }
                    }

                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed.";
                    lbnMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                lbnMessage.Text = "Please insert data source before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }

        private void GetNameKensa(string _pathexcel, string _sheetName)
        {
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel, ReadOnly: false);

            // go to sheet 
            Microsoft.Office.Interop.Excel.Sheets _excelsheets = wb.Worksheets;

            //get contains excel
            Microsoft.Office.Interop.Excel.Worksheet _excelworksheet = _excelsheets.get_Item(_sheetName.Replace("$", "").ToString());
            _excel.DisplayAlerts = false; // disable alert message box 


            #region  Get name of kensa with cell

            if (_excelworksheet.Range[txtCell.Text].get_Value(misValue) != null)
            {
                #region Create and edit file name with regex
                Regex pattern = new Regex("[;,\t\r  < > ? *]");
                string strName = _excelworksheet.Range[txtCell.Text].get_Value(misValue).Split(':')[1];
                strName = pattern.Replace(strName, "");

                string[] _liststr = _pathexcel.Split('\\');
                lbnMessage.Text = strName + " is pending . . . ";
                lbnQuantity.Text = _flagcount + " / " + treeView1.Nodes.Count.ToString();
                #endregion

                #region"save into new folder"
                string _newPath = _pathexcel.Replace(_liststr[_liststr.Length - 1].ToString(), "New Name");
                if (!System.IO.Directory.Exists(_newPath))
                {
                    System.IO.Directory.CreateDirectory(_newPath);
                }
                #endregion

                #region Create new file excel

                //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = _excel.Workbooks.Add(1);
                Workbook newWorkbook = _excel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet newWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)newWorkbook.Worksheets[1]);

                _excelworksheet.Copy(newWorksheet);

                #endregion

                string _newpath = _pathexcel.Replace(_liststr[_liststr.Length - 1].ToString(), "New Name\\" + strName + ".xls");

                if (_newpath.Substring(_newpath.Length - 4) == "xlsx")
                {
                    //wb.saveas
                    newWorkbook.SaveAs(_newpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                }
                else
                {
                    newWorkbook.SaveAs(_newpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing,
                       Type.Missing, false, misValue, XlSaveAsAccessMode.xlExclusive,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                newWorkbook.Close();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(newWorkbook);

            }
            #endregion

            wb.Save();

            wb.Close();
            _workbook.Close();


            _excel.DisplayAlerts = true;
            _excel.Quit();

            // release all
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_workbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
            _excel = null;
            GC.Collect();

            // kill all process name Excel
            Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
            if (!excelProcess.CloseMainWindow())
            {
                excelProcess.Kill();
            }
        }

        private void GetNameCheckSheet(string _pathexcel)
        {
            // Create and edit file name with regex
            Regex pattern = new Regex("[;,\t\r  < > ? *]");
            string[] _liststr = _pathexcel.Split('\\');
            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_pathexcel, ReadOnly: false);

            // go to sheet 
            Microsoft.Office.Interop.Excel.Sheets _excelsheets = wb.Worksheets;




            //Microsoft.Office.Interop.Excel.Worksheet _excelworksheet =
            foreach (Worksheets _excelworksheet in wb.Worksheets)
            {
                //get contains excel
                Microsoft.Office.Interop.Excel.Worksheet _excelws = _excelsheets.get_Item(_excelworksheet.ToString().Replace("$", "").ToString());
                _excel.DisplayAlerts = false; // disable alert message box 

                string strName = _excelws.Range[txtCell.Text].get_Value(misValue).Split(':')[1];
                strName = pattern.Replace(strName, "");

                var _newbook = _excel.Workbooks.Add(1);
                _excelworksheet.Copy(_newbook.Sheets[1]);

                string _newpath = _pathexcel.Replace(_liststr[_liststr.Length - 1].ToString(), "New Name\\.xls");




                // save file excel with get type of file
                if (_newpath.Substring(_newpath.Length - 4) == "xlsx")
                {
                    wb.SaveAs(_newpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                }
                else
                {
                    wb.SaveAs(_newpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing,
                       Type.Missing, false, misValue, XlSaveAsAccessMode.xlExclusive,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }

            wb.Save();

            wb.Close();
            _workbook.Close();
            _excel.DisplayAlerts = true;
            _excel.Quit();

            // release all
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelsheets);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_workbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel);
            _excel = null;
            GC.Collect();

            // kill all process name Excel
            Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
            if (!excelProcess.CloseMainWindow())
            {
                excelProcess.Kill();
            }


        }


        private void CreateLog(string _ex)
        {
            StreamWriter _writer = new StreamWriter("error_log.txt");
            _writer.Write(DateTime.Now.ToString() + _ex);
            _writer.Close();
        }


        private void btnGetPathExcelforKenSa(object sender, EventArgs e)
        {
            int _Countdtb = _dtbExcel.Rows.Count;

            //お客様はベックマンの図面を探したいんです。

            // Check Name by Treeview --> Get Path
            if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
            {
                try
                {
                    for (int i = 0; i < _Countdtb; i++)
                    {

                        string _RowExecute = _dtbExcel.Rows[i]["Name"].ToString();
                        foreach (TreeNode _node in treeView1.Nodes)
                        {
                            if (_node.Text.Contains(_RowExecute))
                            {
                                _dtbExcel.Rows[i]["Path"] = _node.Tag.ToString();
                                break;
                            }
                            else // if not found 
                            {
                                if (_RowExecute.Length > 8)
                                {
                                    string _start = _RowExecute.Remove(8);
                                    string _end = _RowExecute.Substring(_RowExecute.Length - 2, 2);
                                    string _lastchar = _RowExecute.Substring(_RowExecute.Length - 1, 1);
                                    if ((_node.Text.Contains(_start) && _node.Text.Contains(_end) && Char.IsLetter(_lastchar[0])) ||
                                        _node.Text.Contains(_RowExecute.Remove(6)) ||
                                    (_node.Text.Contains(_start) && Char.IsLetter(_lastchar[0])))
                                    {
                                        _dtbExcel.Rows[i]["Path"] = _node.Tag.ToString();
                                    }
                                }
                                else
                                {
                                    //--> check ban ghep
                                    if ((_node.Text == _RowExecute) || _node.Text.Remove(6).Contains(_RowExecute))
                                    {
                                        _dtbExcel.Rows[i]["Path"] = _node.Tag.ToString();
                                    }
                                }
                            }
                        }
                    }
                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;


                    // Check list not founds
                    System.Data.DataTable _dtbError = new System.Data.DataTable();
                    _dtbError.Columns.Add("Name");
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                        {
                            _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
                        }

                    }
                    if (_dtbError.Rows.Count > 0)
                    {
                        dtgError.DataSource = _dtbError;
                    }
                    lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
                    _dtbError = new System.Data.DataTable();
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed";
                    lbnMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                lbnMessage.Text = "Please insert data source before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }

        #region "古いコード

        //else
        //{
        //    // Check Name by Treeview --> Get Path
        //    if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
        //    {
        //        try
        //        {
        //            for (int i = 0; i < _Countdtb; i++)
        //            {
        //                foreach (TreeNode _node in treeView1.Nodes)
        //                {
        //                    string _nodetext = _node.Text.Replace("-", "").Replace(" ", "").Trim().ToUpper();

        //                    if (_nodetext.Contains(_dtbExcel.Rows[i]["Name"].ToString().Replace("-", "").Replace(" ", "").Trim().ToUpper()))
        //                    {
        //                        _dtbExcel.Rows[i]["Path"] = _node.Tag.ToString();
        //                    }
        //                }
        //            }
        //            lbnMessage.Text = "Get path excel successful";
        //            lbnMessage.ForeColor = Color.Blue;


        //            // Check list not founds
        //            System.Data.DataTable _dtbError = new System.Data.DataTable();
        //            _dtbError.Columns.Add("Name");
        //            for (int i = 0; i < _Countdtb; i++)
        //            {
        //                if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
        //                {
        //                    _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
        //                }

        //            }
        //            if (_dtbError.Rows.Count > 0)
        //            {
        //                dtgError.DataSource = _dtbError;
        //            }
        //            lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
        //            _dtbError = new System.Data.DataTable();
        //        }
        //        catch (Exception ex)
        //        {
        //            lbnMessage.Text = "Get path excel failed";
        //            lbnMessage.ForeColor = Color.Red;
        //        }
        //    }
        //    else
        //    {
        //        lbnMessage.Text = "Please insert data source before.";
        //        lbnMessage.ForeColor = Color.Blue;
        //    }
        //}
        #endregion


        // just get name of file kensa or check sheet.
        // if have version name --> cut version name and save as new file
        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int _count = _dtbTreeView.Rows.Count;
                if (_count > 0)
                {

                    // delete file type and " "
                    for (int i = 0; i < _count; i++)
                    {
                        string _strPath = _dtbTreeView.Rows[i]["Path"].ToString();
                        string _Name = _dtbTreeView.Rows[i]["Name"].ToString();

                        _Name.Replace(" ", "");

                        if (_Name.Substring(_Name.Length - 4, 4) == ".xls")
                            _Name = _Name.Remove(_Name.Length - 4, 4);
                        else
                            _Name = _Name.Remove(_Name.Length - 5);
                        // update file name and path item
                        _dtbTreeView.Rows[i]["Path"].ToString().Replace(_dtbTreeView.Rows[i]["Name"].ToString(), _Name);
                        _dtbTreeView.Rows[i]["Name"] = _Name;

                    }


                    for (int i = 0; i < _count; i++)
                    {
                        string _strPath = _dtbTreeView.Rows[i]["Path"].ToString();
                        string _Name = _dtbTreeView.Rows[i]["Name"].ToString();

                        //create new file name
                        // if file name as the same MF618800 or MF6188F1
                        if (_Name.Length < 10 && _Name.Length > 6)
                        {
                            if (Char.IsNumber(_strPath.Substring(_strPath.Length - 1, 1)[0]))
                            {
                                //create new name file 
                                String _newName = _Name.Remove(6);

                                //replace old filename with new filename
                                //String _newPath = _dtbTreeView.Rows[i]["Path"].ToString().Replace(_Name, _newName);

                                // rename excel file
                                System.IO.File.Move(_strPath, _strPath.Replace(_Name, _newName));
                            }
                        }
                        else // if file name as the same MF4882F1-4-B or MF4789F1-3-B&Ghép or B01611F1-1
                        {
                            List<String> _ListSplit = _Name.Split('-').ToList();
                            if (_ListSplit.Count > 2)
                            {
                                String _newName = _ListSplit[0] + _ListSplit[_ListSplit.Count - 1];
                                // replace name
                            }
                            else
                            {
                                string _newName = _Name.Remove(6);
                            }

                        }
                    }
                }


            }
            catch (Exception)
            {


            }
        }

        //just print excel
        private void btnJustPrint_Click(object sender, EventArgs e)
        {
            int _countExcel = _dtbExcel.Rows.Count;
            lbnQuantity.Text = String.Empty;

            try
            {
                for (int i = 0; i < _countExcel; i++)
                {
                    if (!string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                    {
                        if (!_flagbtnHino)
                            PrintWithSheetName(_dtbExcel.Rows[i]["Path"].ToString(), _dtbExcel.Rows[i]["Name"].ToString());
                        else
                            PrintWithSheetName(_dtbExcel.Rows[i]["Path"].ToString(), _dtbExcel.Rows[i]["ID"].ToString());

                        lbnMessage.Text = _dtbExcel.Rows[i]["Name"].ToString() + " is pending . . .";
                        lbnQuantity.Text = (i + 1).ToString() + " / " + _countExcel.ToString();
                        lbnMessage.ForeColor = Color.Blue;
                    }
                }
                _dtbExcel = new System.Data.DataTable();
                lbnMessage.Text = "Auto edit and print successful";
                lbnMessage.ForeColor = Color.Blue;
                this.BringToFront();
            }
            catch (Exception)
            {
                _dtbExcel = new System.Data.DataTable();
                lbnMessage.Text = "Auto edit and print fail, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
        }

        private void getPathOfHinnoItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool _flag = false;
            int _Countdtb = _dtbExcel.Rows.Count;
            // Check Name by Treeview --> Get Path
            if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
            {
                try
                {
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        foreach (TreeNode _node in treeView1.Nodes)
                        {
                            foreach (TreeNode _childNode in _node.Nodes)
                            {
                                string _str = _childNode.Text.Replace("'", "").Replace("$", "").Trim();
                                //_str =  _childNode.Text.Replace("$", "");
                                //select =
                                if ((_dtbExcel.Rows[i]["Name"].ToString() == _str) || ((_dtbExcel.Rows[i]["Name"].ToString() + " Laser") == _str) || ((_dtbExcel.Rows[i]["Name"].ToString() + " Laze") == _str))
                                {
                                    _dtbExcel.Rows[i]["Path"] = _childNode.Tag.ToString();
                                    _flag = true;
                                    break;
                                }
                            }
                            if (_flag) break;

                        }
                        _flag = false;
                    }
                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;


                    // Check list not founds
                    System.Data.DataTable _dtbError = new System.Data.DataTable();
                    _dtbError.Columns.Add("Name");
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                        {
                            _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
                        }
                    }
                    if (_dtbError.Rows.Count > 0)
                    {
                        dtgError.DataSource = _dtbError;
                    }
                    lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
                    _dtbError = new System.Data.DataTable();
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed";
                    lbnMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                lbnMessage.Text = "Please insert data source before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }

        private void HinoGetPathNameKensa_Click(object sender, EventArgs e)
        {
            int _Countdtb = _dtbExcel.Rows.Count;
            // Check Name by Treeview --> Get Path
            if ((treeView1.Nodes.Count > 0) && (_Countdtb > 0))
            {
                try
                {
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        foreach (TreeNode _node in treeView1.Nodes)
                        {
                            string _nodetext = _node.Text.Replace("-", "").Replace(" ", "").Trim().ToUpper();

                            if (_nodetext.Contains(_dtbExcel.Rows[i]["Name"].ToString().Replace("-", "").Replace(" ", "").Trim().ToUpper()))
                            {
                                _dtbExcel.Rows[i]["Path"] = _node.Tag.ToString();
                            }
                            else // 
                            {

                            }
                        }
                    }
                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;


                    // Check list not founds
                    System.Data.DataTable _dtbError = new System.Data.DataTable();
                    _dtbError.Columns.Add("Name");
                    for (int i = 0; i < _Countdtb; i++)
                    {
                        if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                        {
                            _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
                        }

                    }
                    if (_dtbError.Rows.Count > 0)
                    {
                        dtgError.DataSource = _dtbError;
                    }
                    lbnQuantity.Text = (_Countdtb - _dtbError.Rows.Count).ToString() + " / " + _Countdtb.ToString();
                    _dtbError = new System.Data.DataTable();
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed";
                    lbnMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                lbnMessage.Text = "Please insert data source before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void ckbEditBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbEditBox.Checked == true)
            {
                pnlEditName.Visible = true;
                txtEdit.Text = "";
                txtEdit.Focus();
            }
            else
            {
                pnlEditName.Visible = false;
            }

        }

        private void rdbCompany_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnEnglishLanguage_Click(object sender, EventArgs e)
        {
            btnEnglishLanguage.Checked = true;
            SetLanguage("en-EN");
        }

        private void btnVNLanguage_Click(object sender, EventArgs e)
        {
            try
            {
                btnVNLanguage.Checked = true;
                btnEnglishLanguage.Checked = false;
                SetLanguage("vi-VN");
            }
            catch (Exception)
            {
                lbnMessage.Text = "Change language fail, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
        }

        private void btnEnglishLanguage_Click_1(object sender, EventArgs e)
        {
            try
            {
                btnEnglishLanguage.Checked = true;
                btnVNLanguage.Checked = false;
                SetLanguage("en-US");
            }
            catch (Exception)
            {
                lbnMessage.Text = "Change language fail, please try later";
                lbnMessage.ForeColor = Color.Red;
            }
        }

        private void btnBKM_Click(object sender, EventArgs e)
        {
            // set btn BeckMan is choosed
            SetDefaultButton(true, false, false);
        }

        private void btnRRC_Click(object sender, EventArgs e)
        {
            SetDefaultButton(false, true, false);
        }

        private void btnHino_Click(object sender, EventArgs e)
        {
            SetDefaultButton(false, false, true);
        }

        private void importListDefineHinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = false;
                if (_open.ShowDialog() == DialogResult.OK)
                {
                    _dtbListHino = GetDataTable(_open.FileName.ToString(), _open.SafeFileName.ToString());
                    dtgListSourceHino.DataSource = _dtbListHino;
                    lbnMessage.Text = "Import List define of Hino successfull.";
                    lbnQuantity.Text = "";
                }
            }
            catch (Exception)
            {
                lbnText.Text = "Import list define fail.";
            }
        }


        private void AutoCompareData(System.Data.DataTable _dtbSource, System.Data.DataTable _dtbExcute)
        {
            //add new column into datatable
            System.Data.DataColumn _column = new DataColumn("ID", typeof(System.String));
            _dtbExcute.Columns.Add(_column);
            _column.SetOrdinal(0);

            int _countSource = _dtbSource.Rows.Count;
            int _countExe = _dtbExcute.Rows.Count;
            try
            {
                if (_countSource > 0 && _countExe > 0)
                {
                    for (int i = 0; i < _countSource; i++)
                    {
                        for (int j = 0; j < _countExe; j++)
                        {
                            if (_dtbExcute.Rows[j]["Name"].ToString() == _dtbSource.Rows[i]["Name"].ToString())
                            {
                                _dtbExcute.Rows[j]["ID"] = _dtbSource.Rows[i]["ID"].ToString();
                            }
                        }
                    }
                }

            }
            catch (Exception)
            {
                lbnText.Text = "Compare fail. Please try late.";
            }
        }

        private void btnEditAndPrint_Click(object sender, EventArgs e)
        {

        }

        private void btnClearListHino_Click(object sender, EventArgs e)
        {
            dtgListSourceHino.DataSource = new System.Data.DataTable();
        }

        private void btnFixNameCheckSheet_Click(object sender, EventArgs e)
        {
            // Check Name by Treeview --> Get Path
            if (treeView1.Nodes.Count > 0 && !string.IsNullOrEmpty(txtCell.Text))
            {
                try
                {
                    foreach (TreeNode _node in treeView1.Nodes)
                    {
                        _flagcount++;

                        //GetNameKensa(_childNode.Tag.ToString(), _childNode.Text);
                        AutoSaveEachSheet(_node.Tag.ToString());

                        // kill all process name Excel
                        Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
                        if (!excelProcess.CloseMainWindow())
                        {
                            excelProcess.Kill();
                        }
                    }

                    lbnMessage.Text = "Get path excel successful";
                    lbnMessage.ForeColor = Color.Blue;
                }
                catch (Exception)
                {
                    lbnMessage.Text = "Get path excel failed.";
                    lbnMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                lbnMessage.Text = "Please insert data source before.";
                lbnMessage.ForeColor = Color.Blue;
            }
        }

        private void AutoSaveEachSheet(string _path)
        {


            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();

            //disable alert in excel
            _excel.DisplayAlerts = false;

            //create workbook
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_path, ReadOnly: false);



            foreach (Worksheet _sheet in wb.Worksheets)
            {
                var newbook = _excel.Workbooks.Add(1);
                _sheet.Copy(newbook.Sheets[1]);


                #region Create and edit file name with regex
                Regex pattern = new Regex("[;,\t\r  < > ? *]");
                string strName = "";

                try
                {
                    //get name of checksheet E3 or E4 cell
                    if (_sheet.Range["E3"].get_Value(misValue) == null)
                        strName = _sheet.Range["E4"].get_Value(misValue).Split(':')[1];
                    else if (_sheet.Range["E4"].get_Value(misValue) == null)
                        strName = _sheet.Range["E3"].get_Value(misValue).Split(':')[1];
                    else
                    {
                        //add list error
                        System.Data.DataTable _dtb = new System.Data.DataTable();
                        DataColumn _col = new DataColumn();
                        _col.ColumnName = "List Error";
                        _dtb.Columns.Add(_col);
                        _dtb.Rows.Add(_path);
                        dtgError.DataSource = _dtb;
                    }


                    strName = pattern.Replace(strName, "");
                }
                catch (Exception)
                {
                }

                string[] _liststr = _path.Split('\\');
                lbnQuantity.Text = _flagcount + " / " + treeView1.Nodes.Count.ToString();
                #endregion

                #region"save into new folder"
                string _newPath = _path.Replace(_liststr[_liststr.Length - 1].ToString(), "New Name");
                if (!System.IO.Directory.Exists(_newPath))
                {
                    System.IO.Directory.CreateDirectory(_newPath);
                }
                #endregion

                newbook.SaveAs(_newPath + "\\" + _sheet.Name);
                newbook.Close();

                _excel.DisplayAlerts = true;


            }



        }



        private void AutoSaveEachSheetKenSa(string _path)
        {


            object misValue = System.Reflection.Missing.Value;

            // open file excel
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();

            //disable alert in excel
            _excel.DisplayAlerts = false;

            //create workbook
            var _workbook = _excel.Workbooks;
            Workbook wb = _workbook.Open(_path, ReadOnly: false);



            foreach (Worksheet _sheet in wb.Worksheets)
            {
                var newbook = _excel.Workbooks.Add(1);
                _sheet.Copy(newbook.Sheets[1]);


                #region Create and edit file name with regex
                Regex pattern = new Regex("[;,\t\r  < > ? *]");
                string strName = "";

                try
                {
                    //get name of checksheet E3 or E4 cell
                    //if (_sheet.Range["A3"].get_Value(misValue) == null)
                    //    strName = _sheet.Range["A2"].get_Value(misValue).Split(':')[1];
                    //else if (_sheet.Range["A2"].get_Value(misValue) == null)
                    //    strName = _sheet.Range["A3"].get_Value(misValue).Split(':')[1];
                    //else
                    //{
                    //    //add list error
                    //    System.Data.DataTable _dtb = new System.Data.DataTable();
                    //    DataColumn _col = new DataColumn();
                    //    _col.ColumnName = "List Error";
                    //    _dtb.Columns.Add(_col);
                    //    _dtb.Rows.Add(_path);
                    //    dtgError.DataSource = _dtb;
                    //}


                    strName = _sheet.Range["A2"].get_Value(misValue).Split(':')[1];

                    strName = pattern.Replace(strName, "");
                }
                catch (Exception)
                {
                }

                string[] _liststr = _path.Split('\\');
                lbnQuantity.Text = _flagcount + " / " + treeView1.Nodes.Count.ToString();
                #endregion

                #region"save into new folder"
                string _newPath = _path.Replace(_liststr[_liststr.Length - 1].ToString(), "New Name");
                if (!System.IO.Directory.Exists(_newPath))
                {
                    System.IO.Directory.CreateDirectory(_newPath);
                }
                #endregion

                newbook.SaveAs(_newPath + "\\" + strName);
                newbook.Close();

                _excel.DisplayAlerts = true;


            }



        }


        #region "Print pdf file"
        public static bool Print(string file, string printer)
        {
            try
            {
                Process.Start(
                   Registry.LocalMachine.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows\CurrentVersion" +
                        @"\App Paths\AcroRd32.exe").GetValue("").ToString(),
                   string.Format("/h /t \"{0}\" \"{1}\"", file, printer));
                return true;
            }
            catch { }
            return false;
        }
        #endregion

        private void btnPdfDataSource_Click(object sender, EventArgs e)
        {

            if (_dtbTreeView.Columns.Count < 1)
            {
                _dtbTreeView.Columns.Add("Name");
                _dtbTreeView.Columns.Add("Path");
            }
            else
            {
                treeView1.Nodes.Clear();
            }
            try
            {

                // Open File Dialog
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = true;

                if (_open.ShowDialog() == DialogResult.OK)
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    _listFileName = _open.SafeFileNames;

                    //define variable


                    //add list to tree view
                    int _countList = _listFileName.Count();
                    for (int i = 0; i < _countList; i++)
                    {
                        string _Name = _open.SafeFileNames[i].ToString().Trim();
                        string _Path = _open.FileNames[i].ToString().Trim();

                        //iranai
                        #region Check trung
                        ////Check trung
                        //foreach (TreeNode _Node in treeView1.Nodes)
                        //{
                        //    if (_Node.Name == _Name)
                        //        break;
                        //}
                        #endregion
                        AddpdfintoTreeView(_Path, _Name);
                        //add data to temp table

                        _dtbTreeView.Rows.Add(_Name, _Path);
                    }

                    // go to treeview tab
                    lbnQuantity.Text = _countList.ToString();
                    //System.IO.File.Move(@"G:\Work\Tool Data\List kensa\B05117F1-A & Ghép.xls", @"G:\Work\Tool Data\List kensa\B05117F1-A & Ghep.xls");
                    lbnMessage.Text = "Import data source successful";
                    lbnMessage.ForeColor = Color.Blue;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Import fail...");
            }
        }

        private void btnPrintDrawing_Click(object sender, EventArgs e)
        {
            PrintDrawingPdf(@"C:\Users\チュン\Dropbox\Nhap bieu cong trinh\Beckman & TMSC\Ban ve kiem tra BC\Inox 1.5\B01611.pdf");
        }

        public void PrintDrawingPdf(string _pathpdf)
        {
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo()
            {
                CreateNoWindow = true,
                Verb = "print",
                FileName = _pathpdf //put the correct path here
            };
            p.Start();
        }


        //2016/07/08_HonC_ Get printer
        public string GetPrinter()
        {
            PrinterSettings _setting = new PrinterSettings();
            foreach (string _printer in PrinterSettings.InstalledPrinters)
            {
                _setting.PrinterName = _printer;
                if (_setting.IsDefaultPrinter)
                {

                    return _printer;
                }
            }
            return string.Empty;
        }

        //2016/07/08_HonC_Get Tray of Printer
        public List<string> getTrayPrinter()
        {
            _List = new List<string>();
            //Get printer
            PrinterSettings _setting = new PrinterSettings();
            foreach (var _printer in PrinterSettings.InstalledPrinters)
            {
                //if printer is default printer
                if (_setting.IsDefaultPrinter)
                {

                    foreach (System.Drawing.Printing.PaperSource item in _setting.DefaultPageSettings.PrinterSettings.PaperSources)
                    {
                        _List.Add(item.SourceName.ToString());
                    }
                    break;
                }
                else
                    continue;
            }
            return _List;
        }


        //2016/07/08_HonC_Set tray for Printer
        public void SetTrayPrinter(int _indexPaperSource)
        {
            //get list tray printer
            List<string> _List = getTrayPrinter();

            PrinterSettings _setting = new PrinterSettings();


        }

        private void testPrinterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _List = getTrayPrinter();
                if (_List.Count > 0)
                {
                    //get list tray printer

                    System.Data.DataTable _dtb = new System.Data.DataTable();
                    _dtb.Columns.Add("ID");
                    _dtb.Columns.Add("Value");

                    // if list tray printer not null -> add to combobox
                    for (int i = 0; i < _List.Count; i++)
                    {
                        DataRow _row = _dtb.NewRow();
                        _row["Value"] = _List[i].ToString();     //.Replace("PaperSource", "");
                        _row["ID"] = i.ToString();
                        _dtb.Rows.Add(_row);
                    }
                    // delete first row
                    _dtb.Rows[0].Delete();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void cmbPageSource_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void inputExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Convert Excel to pdf tempt
            string _ahihi = LibPrintExcel.ExcelToPdf(@"C:\Users\HonC\Desktop\EDI受注一覧.xls", @"C:\Users\HonC\Desktop\Untitled.pdf");

            //Change setting

            //   string _ppSource = _List[int.Parse(cmbPageSource.SelectedValue.ToString())].ToString();
            //LibPrintExcel.ReadPdfAndPrint(_ahihi, _ppSource);

        }


        //2016/07/11_HonC print document with Path and TrayName of Printer
        public void PrintExcelwithPaathAndTrayName(string _Path, string _TrayName)
        {
            const int WIN_1252_CP = 1252;
            #region "Use stream Reader and Print Document for print excel"

            _StrReader = new StreamReader(_Path, Encoding.GetEncoding(WIN_1252_CP), false);
            PrintDocument _PrDocument = new PrintDocument();


            printFont = new System.Drawing.Font("Times New Roman", 10);
            _PrDocument.PrintPage += new PrintPageEventHandler(pd_PrintPage);





            foreach (PaperSource _pSource in _PrDocument.PrinterSettings.PaperSources)
            {
                if (_pSource.SourceName.ToUpper() == _TrayName.ToUpper())
                {
                    _PrDocument.DefaultPageSettings.PaperSource = _pSource;
                    break;
                }
            }

            _PrDocument.Print();
            #endregion


        }


        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            float linesPerPage = 0;
            float yPos = 0;
            int count = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            String line = null;

            // Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);

            // Iterate over the file, printing each line.
            while (count < linesPerPage &&
               ((line = _StrReader.ReadLine()) != null))
            {
                yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));
                ev.Graphics.DrawString(line, printFont, Brushes.Black,
                   leftMargin, yPos, new StringFormat());
                count++;
            }

            // If more lines exist, print another page.
            if (line != null)
                ev.HasMorePages = true;
            else
                ev.HasMorePages = false;
        }

        //read bitmap 
        private void PrintExcelWithBitMap(string _ppSource)
        {
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            Graphics graphics = Graphics.FromImage(printscreen as Image);

            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            printscreen.Save(@"C:\Users\HonC\Desktop\TestTray.xlsx", ImageFormat.Jpeg);
            bmp = printscreen;
            PrintDocument pd = new PrintDocument();
            pd.OriginAtMargins = true;
            foreach (PaperSource _pSource in pd.PrinterSettings.PaperSources)
            {
                if (_pSource.SourceName.ToString().ToUpper() == _ppSource)
                {
                    pd.DefaultPageSettings.PaperSource = _pSource;
                    break;
                }
            }
            pd.DefaultPageSettings.Landscape = true;
            pd.PrintPage += this.Doc_PrintPage;
            pd.Print();
        }

        private void PrintExcelWithBitMap(Bitmap _bm, string _tray)
        {
            bmp = _bm;
            PrintDocument pd = new PrintDocument();
            pd.OriginAtMargins = true;

            //foreach (PaperSource _pSource in pd.PrinterSettings.PaperSources)
            //{
            //    if (_pSource.SourceName.ToString().ToUpper() == _tray)
            //    {
            //        pd.DefaultPageSettings.PaperSource = _pSource;
            //        break;
            //    }
            //}

            pd.PrinterSettings.Duplex = Duplex.Simplex;
            pd.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("PaperA4", 840, 1180);

            pd.PrintPage += this.Doc_PrintPage;
            pd.Print();
        }

        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            double cmToUnits = 100 / 2.54;
            e.Graphics.DrawImage(bmp, 0, 0, (float)(27 * cmToUnits), (float)(18 * cmToUnits));
        }

        private void btnCustomPrinter_Click(object sender, EventArgs e)
        {
            if (!File.Exists(LibPrintExcel._DEFAULTPDFTEMPPATH))
            {
                //Create TempPDF file
                LibPrintExcel._DEFAULTPDFTEMPPATH = LibPrintExcel.AutoCreateTempPdf();
            }

            wfCustomPrinter_UI _LoadPrinter = new wfCustomPrinter_UI();
            _LoadPrinter.ShowDialog();
        }

        private void printCustomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Check DEFAULT PDF TEMP PATH
            if (!File.Exists(LibPrintExcel._DEFAULTPDFTEMPPATH))
            {
                //Create TempPDF file
                LibPrintExcel._DEFAULTPDFTEMPPATH = LibPrintExcel.AutoCreateTempPdf();
            }

            if (LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.FromPage > 0)
            {
                lbnMessage.Text = string.Empty;
                lbnMessage.ForeColor = Color.Blue;

                // Default settings and custom setting together use global variable
                int _dtbExcelCount = _dtbExcel.Rows.Count;
                for (int i = 0; i < _dtbExcelCount; i++)
                {
                    string _tempPath = _dtbExcel.Rows[i]["Path"].ToString();
                    //Convert to PDf temp
                    LibPrintExcel.ExcelToPdf(_tempPath, LibPrintExcel._DEFAULTPDFTEMPPATH);


                    //Print pdf File with Default Settings
                    LibPrintExcel.PrintPdfWithPrinterSettings(LibPrintExcel._DEFAULTPDFTEMPPATH, LibPrintExcel._DEFAULTSETTINGS);
                    lbnMessage.Text = _dtbExcel.Rows[i]["Name"].ToString() + " is Printing ...";
                }

                lbnMessage.Text = "Print Success.";
            }
            else
            {
                lbnMessage.Text = "Please set printer setting before Insert.";
                lbnMessage.ForeColor = Color.Red;
            }
            LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.FromPage = 0;
        }


        /// <summary>
        /// 2016/07/19_HonC
        /// Print list Excel 
        /// </summary>
        /// <param name="_Dtg">DataTable with Name and Path Columns</param>
        public void PrintListExcel(System.Data.DataTable _Dtg)
        {
            int _DtgCount = _Dtg.Rows.Count;
            for (int i = 0; i < _DtgCount; i++)
            {
                LibPrintExcel.ExcelToPdf(_Dtg.Rows[i]["Path"].ToString(), LibPrintExcel._DEFAULTPDFTEMPPATH);
                LibPrintExcel.PrintPdfWithPrinterSettings(LibPrintExcel._DEFAULTPDFTEMPPATH, LibPrintExcel._DEFAULTSETTINGS);
                LibPrintExcel._DEFAULTPDFTEMPPATH = string.Empty;
            }
        }



        private void mnsImportCheckSheet_Click(object sender, EventArgs e)
        {

            if (_dtbTreeView.Columns.Count < 1)
            {
                _dtbTreeView.Columns.Add("Name");
                _dtbTreeView.Columns.Add("Path");
            }
            else
            {
                treeView1.Nodes.Clear();
            }

            try
            {
                // Open File Dialog
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = true;

                if (_open.ShowDialog() == DialogResult.OK)
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    _listFileName = _open.SafeFileNames;

                    ///2016/07/20 _HonC
                    ///Add Neame of Checksheet to TreeView
                    int _countList = _listFileName.Count();

                    //Define DataTable for SearchEngine
                    _DataSourcee = new System.Data.DataTable();
                    _DataSourcee.Columns.Add("Name");
                    _DataSourcee.Columns.Add("Path");
                    for (int i = 0; i < _countList; i++)
                    {
                        DataRow _newR = _DataSourcee.NewRow();
                        _newR["Name"] = _open.SafeFileNames[i].ToString();
                        _newR["Path"] = _open.SafeFileNames[i].ToString();
                        _DataSourcee.Rows.Add(_newR);
                    }

                    //add list to tree view
                    // Add Name as Name of CheckSheet
                    // Add Path as Path of CheckSheet
                    for (int i = 0; i < _countList; i++)
                    {
                        TreeNode _TreeNodeTemp = new TreeNode();
                        _TreeNodeTemp.Text = _open.SafeFileNames[i].ToString();
                        _TreeNodeTemp.Tag = _open.FileNames[i].ToString();
                        treeView1.Nodes.Add(_TreeNodeTemp);
                    }

                    // go to treeview tab
                    tabControl1.SelectedTab = tpListDataSource;
                    lbnQuantity.Text = _countList.ToString();
                    lbnMessage.Text = "Import data source successful";
                    lbnMessage.ForeColor = Color.Blue;

                }
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.ToString());
                lbnMessage.Text = "Import data source failed. ";
                lbnMessage.ForeColor = Color.Red;

            }
            Cursor.Current = Cursors.Default;
        }

        private void importListExcuteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (_flagbtnHino == false || _dtbListHino.Rows.Count > 0)
            {
                _dtbExcel = new System.Data.DataTable();
                dtgListPrint.DataSource = new System.Data.DataTable();
                dtgError.DataSource = new System.Data.DataTable();
                // import file excel -> read excel to datagridview
                try
                {
                    // Open File Dialog
                    OpenFileDialog _open = new OpenFileDialog();
                    _open.Filter = "All Files (*.*)|*.*";
                    _open.FilterIndex = 1;
                    _open.Multiselect = false;

                    if (_open.ShowDialog() == DialogResult.OK)
                    {
                        _listFileName = _open.SafeFileNames;

                        //Get list dataSource
                        _dtbExcel = GetDataTable(_open.FileName.ToString(), _open.SafeFileName.ToString());
                        if ((_dtbExcel != null) && _dtbExcel.Rows.Count > 0)
                            dtgListPrint.DataSource = _dtbExcel;

                        dtgListPrint.Columns["Path"].Width = 500;
                        //Check Name and Path of CheckSeet
                        for (int i = 0; i < _dtbExcel.Rows.Count; i++)
                        {
                            if ((_dtbExcel.Rows[i]["Name"]) == null || _dtbExcel.Rows[i]["Name"].ToString() == "")
                                _dtbExcel.Rows[i].Delete();
                        }
                        _dtbExcel.AcceptChanges();

                        //2016/08/03 _HonC
                        #region "Fix name with Define List BeckMan Extends"
                        string _PathBCDefine = Directory.GetCurrentDirectory() + "DinhNghiaBanGhepBC.xlsx";

                        //if file BC define is NOT exist --> copy file from ... to Current Directory
                        if (!File.Exists(_PathBCDefine))
                        {
                            OpenFileDialog _of = new OpenFileDialog();
                            _of.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
                            _of.Multiselect = false;
                            if (_of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                // copy ListDefineBC to Current Directory
                                System.IO.File.Copy(_of.FileName, _PathBCDefine, true);
                            }
                        }
                        _dtbListHino = GetDataTable(_PathBCDefine, "DinhNghiaBanGhepBC.xlsx");
                        _dtbListHino = LibStub.AutoCompleteDefineExcel(_dtbListHino);
                        // Define list DataSource OK
                        LibStub.CheckExtendsData(_dtbListHino, _dtbExcel);
                        _dtbExcel = LibStub.CheckExtendsData(_dtbListHino, _dtbExcel);
                        dtgListSourceHino.DataSource = _dtbExcel;

                        #endregion

                        // go to List Print Tab
                        tabControl1.SelectedTab = tpListPrint;
                        lbnMessage.Text = "Import list name of check sheet successful";
                        lbnMessage.ForeColor = Color.Blue;
                        lbnQuantity.Text = _dtbExcel.Rows.Count.ToString();
                    }

                }
                catch (Exception ex)
                {
                    lbnMessage.Text = "Import list name of check sheet failed";
                    lbnMessage.ForeColor = Color.Red;
                    lbnMessage.Text = "";
                }
            }
            else
            {
                lbnMessage.Text = "Please insert List Define of Hino first and try it later.";
                lbnQuantity.Text = "";
            }
        }

        /// <summary>
        /// 2016/07/20 _HonC
        /// Find path of CheckSheet file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void getPathOfFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Check for each Tree Node in TreeView
            int _dtbCount = _dtbExcel.Rows.Count;
            string _strTemp = string.Empty;

            //Find path for Checksheet
            for (int i = 0; i < _dtbCount; i++)
            {
                foreach (TreeNode _node in treeView1.Nodes)
                {
                    //define strTemp replace 
                    _strTemp = _node.Text.ToString().Replace(".xlsx", "").Replace(".xls", "");
                    if (_dtbExcel.Rows[i]["Name"].ToString() == _strTemp)
                    {
                        _dtbExcel.Rows[i]["Path"] = _node.Tag;
                        break;
                    }
                    else if ((_strTemp.Length > _dtbExcel.Rows[i]["Name"].ToString().Length) && (_dtbExcel.Rows[i]["Name"].ToString() == _strTemp.Remove(_dtbExcel.Rows[i]["Name"].ToString().Length)))
                    {
                        _dtbExcel.Rows[i]["Path"] = _node.Tag;
                    }
                    else
                        continue;
                }
                tabControl1.SelectedTab = tpListPrint;

            }

            dtgListPrint.DataSource = _dtbExcel;

            System.Data.DataTable _dtbError = new System.Data.DataTable();          //define Dtb Error
            _dtbError.Columns.Add("Name");

            //Check if not found --> 
            int _dtbExcelCount = _dtbExcel.Rows.Count;
            for (int i = 0; i < _dtbExcel.Rows.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(_dtbExcel.Rows[i]["Path"].ToString()))
                {
                    _dtbError.Rows.Add(_dtbExcel.Rows[i]["Name"].ToString());
                    _dtbExcel.Rows[i].Delete();
                }
            }
            dtgError.DataSource = _dtbError;

            lbnMessage.Text = "Get Path Excel Succesfull.";
            lbnQuantity.Text =dtgListPrint.Rows.Count.ToString() +  "/" + lbnQuantity.Text;
        }

        private void ahihiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = false;
                if (_open.ShowDialog() == DialogResult.OK)
                {
                    _dtbListHino = GetDataTable(_open.FileName.ToString(), _open.SafeFileName.ToString());
                    _dtbListHino = LibStub.AutoCompleteDefineExcel(_dtbListHino);
                }
                // Define list DataSource OK
                LibStub.CheckExtendsData(_dtbListHino, _dtbExcel);
                _dtbExcel = LibStub.CheckExtendsData(_dtbListHino, _dtbExcel);
                dtgListSourceHino.DataSource = _dtbExcel;
            }
            catch (Exception)
            {
                lbnText.Text = "Import list define fail.";
            }
        }

        private void printMultiPdfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Open File Dialog
                OpenFileDialog _open = new OpenFileDialog();
                _open.Filter = "All Files (*.*)|*.*";
                _open.FilterIndex = 1;
                _open.Multiselect = true;

                if (_open.ShowDialog() == DialogResult.OK)
                {
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    // txtFilePath.Text = _open.SafeFileName;
                    _listFileName = _open.SafeFileNames;


                    //add list to tree view
                    int _countList = _listFileName.Count();
                    for (int i = 0; i < _countList; i++)
                    {
                        PrintDrawingPdf(_open.FileNames[i].ToString());
                        //GetExcelSheetNames(_open.FileNames[i].ToString(), _open.SafeFileNames[i].ToString());

                        //add data to temp table

                        //_dtbTreeView.Rows.Add(_open.SafeFileNames[i].ToString(), _open.FileNames[i].ToString());
                    }

                    // go to treeview tab
                    lbnQuantity.Text = _countList.ToString();
                    lbnMessage.Text = "Import data source successful";
                    lbnMessage.ForeColor = Color.Blue;

                }
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.ToString());
                lbnMessage.Text = "Import data source failed. ";
                lbnMessage.ForeColor = Color.Red;

            }
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
        }
    }
}

