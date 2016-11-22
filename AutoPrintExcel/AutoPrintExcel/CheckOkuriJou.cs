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
    public partial class CheckOkuriJou : Form
    {
        public CheckOkuriJou()
        {
            InitializeComponent();
        }

        private void btnOpenFileSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog _of = new OpenFileDialog();
            _of.Filter = "All File *.* (";
        }
    }
}
