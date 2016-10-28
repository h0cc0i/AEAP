using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Spire.Pdf;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace AutoPrintExcel
{

    /// <summary>
    /// 2016/07/012_HonC
    /// Function for Print Excel by thread
    /// Parrammeter is Path of Excel
    /// </summary>
    public class LibPrintExcel
    {
        object paramMissing = Type.Missing;
        //2016/07/15_HonC define Default Printer Settings
        public static PrintDialog _DEFAULTSETTINGS = new PrintDialog();

        //2016/07/15_HonC define PdfTemp
        public static string _DEFAULTPDFTEMPPATH = Directory.GetCurrentDirectory() + "ReleaseHonC_tempPDF.pdf";

        //2016/07/19_HonC define FromPage ToPage
        public static int _FROMPAGE = 1;
        public static int _TOPAGE = 1;

        //2016/07/21_HonC define flag use Printer Setting Custorm
        public static bool _USECUSTOMSETTINGS = false;


        public static void PrintProcess(string _Path)
        {

            ProcessStartInfo _info = new ProcessStartInfo();
            _info.Verb = "print";
            _info.FileName = _Path;
            _info.CreateNoWindow = true;
            _info.WindowStyle = ProcessWindowStyle.Hidden;

            Process p = new Process();
            p.StartInfo = _info;
            p.Start();

            p.WaitForInputIdle();
            System.Threading.Thread.Sleep(3000);
            if (false == p.CloseMainWindow())
            {
                p.Kill();
            }
        }

        /// <summary>
        /// 2016/07/14_HonC
        /// </summary>
        /// <param name="_ExcelSource">Param Excel Source file</param>
        /// <param name="_PdfDes"> Param PDf Destiny file</param>
        public static string ExcelToPdf(string _ExcelSource, string _PdfDes)
        {
            Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
            var _workbook = _excel.Workbooks;
            try
            {
                string paramExportFilePath = _PdfDes;
                XlFixedFormatType paramExportFormat = XlFixedFormatType.xlTypePDF;
                XlFixedFormatQuality paramExportQuality = XlFixedFormatQuality.xlQualityStandard;
                bool paramOpenAfterPublish = false;
                bool paramIncludeDocProps = true;
                bool paramIgnorePrintAreas = true;
                object paramFromPage = _FROMPAGE;
                object paramToPage = _TOPAGE;


                Microsoft.Office.Interop.Excel.Workbook _wb = _workbook.Open(_ExcelSource);

                _wb.ExportAsFixedFormat(paramExportFormat,
                        paramExportFilePath, paramExportQuality,
                        paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage,
                        paramToPage, paramOpenAfterPublish,
                        Type.Missing);
                return paramExportFilePath;
            }
            catch (Exception )
            {
                throw;
            }
            finally
            {

                // kill all process name Excel
                Process excelProcess = Process.GetProcessesByName("EXCEL")[0];
                if (!excelProcess.CloseMainWindow())
                {
                    excelProcess.Kill();
                }
            }
        }

        /// <summary>
        /// 2016/07/15_HonC
        /// 1.   Convert from Excel file to pdf
        /// 2.   Change printer settings
        /// 3.   Print pdf FileZ
        /// </summary>
        /// <param name="_Path">Path of Excel File</param>
        /// <param name="_PPSource"> Source pdf file</param>
        public static void ReadPdfAndPrint(string _Path, string _PPSource)
        {
            PdfDocument doc = new PdfDocument();

            doc.LoadFromFile(_Path);
            doc.PageSettings.Orientation = Spire.Pdf.PdfPageOrientation.Landscape;


            PrintDialog dialogPrint = new PrintDialog();
            dialogPrint.AllowPrintToFile = true;
            dialogPrint.AllowSomePages = true;
            dialogPrint.PrinterSettings.MinimumPage = 1;
            dialogPrint.PrinterSettings.MaximumPage = doc.Pages.Count;
            dialogPrint.PrinterSettings.FromPage = 1;
            dialogPrint.PrinterSettings.ToPage = doc.Pages.Count;
            dialogPrint.PrinterSettings.Copies = 1;
            dialogPrint.PrinterSettings.Duplex = Duplex.Vertical;

            // set print setting FromPage
            doc.PrintFromPage = dialogPrint.PrinterSettings.FromPage;
            // Set print setting ToPage
            doc.PrintToPage = dialogPrint.PrinterSettings.ToPage;
            // Set print setting Copies
            doc.PrintDocument.PrinterSettings.Copies = dialogPrint.PrinterSettings.Copies;
            // Set print setting PrinterName
            doc.PrinterName = dialogPrint.PrinterSettings.PrinterName;
            // Set print setting Duplex
            doc.PrintDocument.PrinterSettings.Duplex = dialogPrint.PrinterSettings.Duplex;
            // Set print setting Paper Source
            foreach (PaperSource _pSource in dialogPrint.PrinterSettings.PaperSources)
            {
                if (_pSource.SourceName.ToString().ToUpper() == _PPSource)
                {
                    doc.PrintDocument.DefaultPageSettings.PaperSource = _pSource;

                    break;
                }
            }


            PrintDocument printDoc = doc.PrintDocument;
            dialogPrint.Document = printDoc;
            printDoc.Print();


        }

        /// <summary>
        /// 2016/07/15_HonC
        /// Get Tray of Default Printer 
        /// Return DataTable
        /// </summary>
        /// <returns></returns>
        public static System.Data.DataTable GetTrayofPrinter()
        {
            //define tempt DataTable
            System.Data.DataTable _dtb = new System.Data.DataTable();
            _dtb.Columns.Add("TrayIndex");
            _dtb.Columns.Add("TrayName");

            //define printer settings
            PrinterSettings _setting = new PrinterSettings();

            //Count Papersources in default printer
            int _CountTray = _setting.DefaultPageSettings.PrinterSettings.PaperSources.Count;

            if (_setting.IsDefaultPrinter)
            {
                // Add to DataTable from sencond item on List PaperSources
                for (int i = 1; i < _CountTray; i++)
                {
                    System.Data.DataRow _row = _dtb.NewRow();
                    _row["TrayIndex"] = i.ToString();
                    _row["TrayName"] = _setting.DefaultPageSettings.PrinterSettings.PaperSources[i].SourceName.ToString();
                    _dtb.Rows.Add(_row);
                }
            }

            return _dtb;
        }

        /// <summary>
        /// 2016/07/15_HonC
        /// Define Default Combobox Duplex for Printer
        /// </summary>
        /// <returns></returns>
        public static System.Data.DataTable DefaultDuplexCombo()
        {
            //Define dataTable Duplex
            System.Data.DataTable _dtb = new System.Data.DataTable();
            _dtb.Columns.Add("DuplexName");
            _dtb.Columns.Add("DuplexID");

            //add row 1 - Print Default Duplex
            System.Data.DataRow _row = _dtb.NewRow();
            _row["DuplexName"] = "In mặc định";
            _row["DuplexID"] = "1";

            //add row 2 - Print Double sides Vertical Duplex
            //System.Data.DataRow _row2 = _dtb.NewRow();
            //_row2["DuplexName"] = "In 2 mặt chiều dọc"; // in 2 mặt và in tử phải sang trái
            //_row2["DuplexID"] = "2";

            //add row 3 - Print Double sides Horizontal Duplex
            System.Data.DataRow _row3 = _dtb.NewRow();
            _row3["DuplexName"] = "In hai mặt"; /// in 2 mặt và in từ trái sang phải
            _row3["DuplexID"] = "3";

            //add row 4 - Print Single sides
            System.Data.DataRow _row4 = _dtb.NewRow();
            _row4["DuplexName"] = "In một mặt";
            _row4["DuplexID"] = "4";

            //add Row for DataTable Printer duplex
            _dtb.Rows.Add(_row);
            //_dtb.Rows.Add(_row2);
            _dtb.Rows.Add(_row3);
            _dtb.Rows.Add(_row4);
            return _dtb;
        }

        /// <summary>
        /// 20146/07/15_HonC
        /// Set printer Setting and return PrinterDialog
        /// buong ngur
        /// </summary>
        /// <param name="_ppSource">string Name of PaperSource </param>
        /// <param name="_duplexName">Duplex Name settings </param>
        /// <returns></returns>
        public static PrintDialog SetPrinterSettings(string _ppSource, Duplex _duplexName)
        {
            //Define PrintDialg Tranfer
            PrintDialog _PrintDg = new PrintDialog();

            //Set Printer setting Dubplex
            _PrintDg.PrinterSettings.Duplex = _duplexName;

            //Set Printer Paper Source 
            foreach (PaperSource _pp in _PrintDg.PrinterSettings.PaperSources)
            {
                if (_pp.SourceName.ToUpper() == _ppSource.ToUpper())
                {
                    _PrintDg.PrinterSettings.DefaultPageSettings.PaperSource = _pp;
                    break;
                }
            }

            //Set Printer settings from Page to Page

            return _PrintDg;
        }

        /// <summary>
        /// Auto print pdf wit Path of PDf and Printer Setting --- use PrinterDialog Tranfer object
        /// </summary>
        /// <param name="_Path">Path of Pdf file</param>
        /// <param name="_PrtDialog">PrintDialog object has settings </param>
        public static void PrintPdfWithPrinterSettings(string _Path, PrintDialog _PrtDialog)
        {
            string _tempNamePrt = _PrtDialog.PrinterSettings.DefaultPageSettings.PaperSource.SourceName;
            PdfDocument _doc = new PdfDocument();
            _doc.LoadFromFile(_Path);

            //set Printer Settings
            _doc.PrintDocument.DefaultPageSettings.PaperSource = _PrtDialog.PrinterSettings.DefaultPageSettings.PaperSource;


            _doc.PrintDocument.PrinterSettings.Duplex = _PrtDialog.PrinterSettings.Duplex;

            _doc.PrintDocument.PrinterSettings.FromPage = _PrtDialog.PrinterSettings.FromPage;

            _doc.PrintDocument.PrinterSettings.ToPage = _PrtDialog.PrinterSettings.ToPage;

            PrintDocument _PrtDoc = _doc.PrintDocument;
            _PrtDialog.Document = _PrtDoc;

            //2016/07/22 _HonC set default paper source for DEFAULSETTINGS
            //LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.dfe = _PrtDialog.PrinterSettings.DefaultPageSettings.PaperSource;
            LibPrintExcel._DEFAULTSETTINGS.PrinterSettings.DefaultPageSettings.PaperSource = _PrtDoc.DefaultPageSettings.PaperSource;
            _PrtDoc.Print();

        }

        /// <summary>
        /// 2016/07/15_HonC
        /// Return Duplex from Index
        /// </summary>
        /// <param name="_IndexDuplex"> Params is index of Duplex</param>
        /// <returns></returns>
        public static Duplex ReturnDuplex(int _IndexDuplex)
        {
            string _Name = string.Empty;
            switch (_IndexDuplex.ToString())
            {
                case "0":   //Set duplex is default
                    {
                        return Duplex.Default;
                    }
                //case "1":   //Set duplex is Doubles Print and Horizontal
                //    {
                //        return Duplex.Horizontal;
                //    }
                case "1":   //Set duplex is Doubles Print and Vertical
                    {
                        return Duplex.Vertical;
                    }
                case "2":   //Set duplex is Simplex
                    {
                        return Duplex.Simplex;
                    }
                default:
                    {
                        break;
                    }
            }
            return Duplex.Default;
        }

        /// <summary>
        /// 2016/07/15_HonC
        /// Create new Pdf and Save in Source
        /// </summary>
        public static string AutoCreateTempPdf()
        {
            PdfDocument _pdf = new PdfDocument();
            _pdf.SaveToFile(Directory.GetCurrentDirectory() + "HonC_tempPDF.pdf");
            return Directory.GetCurrentDirectory() + "HonC_tempPDF.pdf";
        }


        /// <summary>
        /// 2016/07/19 _HonC
        /// Convert List Excel to 1 File PDF
        /// </summary>
        /// <param name="_dtb">DataTable List Excel </param>
        /// <returns></returns>
        public static string ListExceltoTempWorkbook(System.Data.DataTable _dtb)
        {
            PdfDocument[] PDFResult = new PdfDocument[150];

            string _PathPDF = string.Empty;
            int _dtbCount = _dtb.Rows.Count;

            for (int i = 0; i < _dtbCount; i++)
            {
                using (MemoryStream m1 = new MemoryStream())
                {
                    Spire.Xls.Workbook _workbook = new Spire.Xls.Workbook();
                    _workbook.LoadFromFile(_dtb.Rows[i]["Path"].ToString());
                    _workbook.SaveToStream(m1, Spire.Xls.FileFormat.Version2007);

                    PDFResult[i] = new PdfDocument(m1);

                }

            }


            return _PathPDF;
        }

       
    }

}
