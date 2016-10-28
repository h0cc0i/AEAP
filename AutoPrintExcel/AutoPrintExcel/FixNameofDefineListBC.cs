using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoPrintExcel
{
    public partial class FixNameofDefineListBC : Form
    {
        System.Data.DataTable _dtbExcel;
        public string _TempPath;
        public FixNameofDefineListBC()
        {
            InitializeComponent();

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
                lblMsg.Text = "Get Excel and Sheet failed, please try later";
                lblMsg.ForeColor = Color.Red;
            }
            finally // close and dispose connection
            {
                ObjConn.Close();
                ObjConn.Dispose();
            }
            return dtb;
        }

        private void btnChoose_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _oFG = new OpenFileDialog();
                _oFG.Filter = "All Files (*.*)|*.*";
                _oFG.FilterIndex = 1;
                _oFG.Multiselect = false;
                if (_oFG.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtPathListDefine.Text = _oFG.FileName.ToString();
                    _dtbExcel = new System.Data.DataTable();
                    //Get list dataSource
                    _dtbExcel = GetDataTable(_oFG.FileName.ToString(), _oFG.SafeFileName.ToString());
                    if ((_dtbExcel != null) && _dtbExcel.Rows.Count > 0)
                        dtgDefineBC.DataSource = _dtbExcel;
                    // 
                    _TempPath = _oFG.FileName.ToString();
                }
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.ToString();
            }


        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //2016/09/27 create button Save
            // Copy File Define list BC -> destination.
            // BCリストを存在するかどうかチェックする。
            var _comfirmDialog = MessageBox.Show("Bạn có muốn cập nhật file định nghĩa hàng BeckMan ?","Cập nhật file",MessageBoxButtons.YesNo);
            if (_comfirmDialog == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    string _PathBCDefine = Directory.GetCurrentDirectory() + "DinhNghiaBanGhepBC.xlsx";
                    // Copy list Define BC to Destination
                    System.IO.File.Copy(_TempPath, _PathBCDefine, true);
                    MessageBox.Show("Cập nhật file định nghĩa hàng BeckMan thành công. ", "Thông báo");
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                
            }
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
