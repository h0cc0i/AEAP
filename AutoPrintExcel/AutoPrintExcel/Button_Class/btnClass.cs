using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoPrintExcel.Button_Class
{
    public class btnClass
    {
        private string _Name;
        private string _ForceColor;
        private string _BackColor;
        private string _FlagStt;

        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }


        public string ForceColor
        {
            get { return _ForceColor; }
            set { _ForceColor = value; }
        }


        public string BackColor
        {
            get { return _BackColor; }
            set { _BackColor = value; }
        }


        public string FlagStt
        {
            get { return _FlagStt; }
            set { _FlagStt = value; }
        }
    }
}
