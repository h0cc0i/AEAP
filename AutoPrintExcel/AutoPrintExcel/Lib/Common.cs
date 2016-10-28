using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace AutoPrintExcel
{
    public class Common
    {

        public static object IsNull(object nguon, object dich)
        {
            if (nguon == null || string.IsNullOrEmpty(nguon.ToString()) || nguon.GetType().ToString() == "DBNull")
            {
                return dich;
            }
            else
                return nguon;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }

        }

        public static void ExportExcel(DataGridView dtg, string Path)
        {


            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Object misValue = System.Reflection.Missing.Value;

            String strPath;
            FolderBrowserDialog fileBrowser = new FolderBrowserDialog();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            strPath = Path;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlApp.Worksheets[1];
            xlWorkSheet.Select(Type.Missing);

            for (int i = 0; i <= dtg.RowCount - 1; i++)
            {
                for (int j = 0; j <= dtg.Columns.Count - 1; j++)
                {
                    for (int k = 1; k <= dtg.Columns.Count; k++)
                    {
                        xlWorkSheet.Cells[1][k] = dtg.Columns[k - 1].HeaderText;
                        xlWorkSheet.Cells[i + 2][j + 1] = dtg[j, i].Value.ToString();
                    }
                }
            }

            xlWorkSheet.SaveAs(Path + "\\vbexcel.xlsx");
            xlWorkBook.Close();
            xlApp.Quit();

            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);



        }

        //Auto gent code for ID
        public static string GentCode4ID(string _str, int _len)
        {
            int _lenstr = _str.Length;
            if (_lenstr < _len)
            {
                do
                {
                    _str = "0" + _str;
                    _lenstr = _str.Length;
                } while (_lenstr < _len);
                return _str;
            }
            else
                return _str;
        }


        public static void AutoCreateTabPage(TabControl _tab, UserControl _uc, string _NameTabPage, string _TextTabPage)
        {
            //define user control
            _uc.Dock = DockStyle.Fill;

            //flag for check tabpage is avaiable
            bool _isAvaiable = false;
            int _countTabPages = Convert.ToInt32(_tab.TabPages.Count.ToString());
            if ((_uc != null) && _countTabPages > 0)
            {
                //if tabcontrol has avaiable tabpages
                for (int i = 0; i < _countTabPages; i++)
                {
                    //if name of tabpage is avaiable
                    if (_tab.TabPages[i].Text == _TextTabPage)
                    {
                        _isAvaiable = true;
                        break;
                    }
                }
                if (!_isAvaiable)
                {
                    //if not avaiable --> create new tabpage

                    //define new tab pages
                    TabPage _tpNew = new TabPage();
                    _tpNew.Text = _TextTabPage;
                    _tpNew.Name = _NameTabPage;
                    _tpNew.Controls.Add(_uc);
                    _tab.TabPages.Add(_tpNew);
                }
            }
            else
            {
                //if tabcontrol not avaiable tab pages
                //define new tab pages
                TabPage _tpNew = new TabPage();
                _tpNew.Text = _TextTabPage;
                _tpNew.Name = _NameTabPage;
                _tpNew.Controls.Add(_uc);
                _tab.TabPages.Add(_tpNew);
            }
            //go to tab pages had choose
            _tab.SelectTab(_NameTabPage);

        }

        public static void PeformTextboxKeyDown(KeyEventArgs e, System.Windows.Forms.TextBox _TxtPre, System.Windows.Forms.TextBox _TxtNext)
        {
            if (e.KeyCode == Keys.Up)
                _TxtPre.Select();
            else if (e.KeyCode == Keys.Enter)
                _TxtNext.Select();
        }

        /// <summary>
        /// Select All Control in current UserControl/Form
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

        /// <summary>
        /// 2016/10/19 HonC
        /// </summary>
        /// <param name="excelpathfile"> Excel path Source </param>
        /// <returns></returns>
        public static System.Data.DataTable GetDataTable(string excelfile)
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
              "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\"";
            }
            else
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                      "Data Source=" + excelfile + ";Mode=ReadWrite;Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally // close and dispose connection
            {
                ObjConn.Close();
                ObjConn.Dispose();
            }
            return dtb;
        }

        /// <summary>
        /// Sort DataTable by Column asc or dsc
        /// </summary>
        /// <param name="_dtb">DataTable Source </param>
        /// <param name="_ColName">Sort by Column</param>
        /// <param name="_Direction"> </param>
        /// <returns></returns>
        public static System.Data.DataTable SortDataTable(System.Data.DataTable _dtb, string _ColName, string _Direction)
        {
            _dtb.DefaultView.Sort = _ColName + " " + _Direction;
            _dtb = _dtb.DefaultView.ToTable();
            return _dtb;
        }


        /// <summary>
        /// Write Printer Config to text file
        /// </summary>
        /// <param name="_path"> Path of text file</param>
        /// <param name="_line1"> Line 1 </param>
        /// <param name="_line2"> Line 2 </param>
        /// <param name="_line3"> Line 3 </param>
        public static void WriteTextPrinterConfig(string _path, string _line1, string _line2, string _line3)
        {
            if (!File.Exists(_path))
                File.Create(_path).Dispose();
            else
                File.WriteAllText(_path,string.Empty);

            string[] _lines =  { "========== Printer Config ==========","", _line1,_line2,_line3 };
            File.WriteAllLines(_path, _lines);
        }

    }
}
