using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoPrintExcel
{
    public partial class wfListPageSource : Form
    {
        public wfListPageSource()
        {
            InitializeComponent();
        }

        public wfListPageSource(List<string> _Listsource)
        {
            InitializeComponent();
        }

        private void btnGetPageSource_Click(object sender, EventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }
    }
}
