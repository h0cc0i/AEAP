using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.IO;

namespace AutoPrintExcel
{
    public class LibStub
    {
        public static string _DefaultLanguage = "vi-VN";
        //private static DataTable _dtbExtend;
        private static string _tempCheckSheet = string.Empty;

        //2016/08/01 _HonC
        private static string _tempInDex = string.Empty;

        /// <summary>
        /// 2016/07/20 _HonC
        /// Use Binary Search Engine for Search Path of CheckSheet
        /// </summary>
        /// <param name="_ListSource"> List Source is List Array 2d </param>
        /// <param name="_dtbDes"> DataTable Des </param>
        public static DataTable AutoSeachPath(DataTable _dtbDataSource, DataTable _dtbDes)
        {
            int _DesCount = _dtbDes.Rows.Count;
            int _SourceCount = _dtbDataSource.Rows.Count;
            for (int i = 0; i < _DesCount; i++)
            {
                for (int j = 0; j < _dtbDataSource.Rows.Count; j++)
                {
                    if (_dtbDataSource.Rows[i]["Name"].ToString().Trim() == _dtbDes.Rows[i]["Path"].ToString().Trim())
                    {
                        _dtbDataSource.Rows[i]["Path"] = _dtbDes.Rows[i]["Path"].ToString().Trim();
                        break;
                    }
                    else
                        continue;
                }
            }

            return _dtbDataSource;
        }

        /// <summary>
        /// 2016/07/28 _HonC 
        /// Check list BC define for get current Name of CheckSheet
        /// </summary>
        /// <param name="_dtbSource"></param>
        /// <param name="_dtbDes"></param>
        /// <returns></returns>
        public static System.Data.DataTable CheckExtendsData(System.Data.DataTable _dtbSource, System.Data.DataTable _dtbDes)
        {
            int _DesCount = _dtbDes.Rows.Count;
            //2016/08/01 _HonC
            // loop find extendData
            for (int i = 0; i < _DesCount; i++)
            {
                for (int j = 0; j < _dtbSource.Rows.Count; j++)
                {
                    // if is the first check
                    if (string.IsNullOrEmpty(_tempCheckSheet))
                    {
                        // if Name in Source equal Name in Des
                        if (_dtbSource.Rows[j]["Name"].ToString() == _dtbDes.Rows[i]["Name"].ToString())
                        {
                            //Set temp
                            _tempCheckSheet = _dtbDes.Rows[i]["Name"].ToString();


                            //Delete this row in Des
                            _dtbDes.Rows[i].Delete();
                            _dtbDes.AcceptChanges();

                            //add data from Source to Destination
                            DataRow _newR = _dtbDes.NewRow();
                            _newR["Name"] = _dtbSource.Rows[j]["Detail"].ToString();
                            _dtbDes.Rows.Add(_newR);

                        }
                    }
                    else
                    {
                        if (_dtbSource.Rows[j]["Name"].ToString() == _tempCheckSheet)
                        {
                            //add data from Source to Destination
                            DataRow _newR = _dtbDes.NewRow();
                            _newR["Name"] = _dtbSource.Rows[j]["Detail"].ToString();
                            _dtbDes.Rows.Add(_newR);
                        }
                        else
                            _tempCheckSheet = string.Empty;
                    }
                }
            }

            return _dtbDes;
        }


        /// <summary>
        /// 2016/08/02 _HonC
        /// Auto Complete Define Excel
        /// </summary>
        /// <param name="_dtb"></param>
        /// <returns></returns>
        public static DataTable AutoCompleteDefineExcel(DataTable _dtb)
        {
            if (!string.IsNullOrEmpty(_dtb.Rows[0]["Name"].ToString()))
            {
                for (int i = 1; i < _dtb.Rows.Count; i++)
                {
                    if (string.IsNullOrEmpty(_dtb.Rows[i]["Name"].ToString()))
                        _dtb.Rows[i]["Name"] = _dtb.Rows[i - 1]["Name"].ToString();
                }
            }
            return _dtb;
        }

        /// <summary>
        /// 2016/08/04_HonC
        /// Check Default Font incurent conputer
        /// </summary>
        public static void CheckDefaultFont()
        {
            /* IF in current Coputer this Font SketchFlow Print is nos exist
             * Add this font from resource  th*/
            if (!File.Exists(@"C:\Windows\Fonts\Code39.ttf"))
            {
                File.Copy(@"\\Lib\\Fonts\\Code39.ttf", @"C:\Windows\Fonts\Code39.ttf", true);
            }
        }
    }
}
