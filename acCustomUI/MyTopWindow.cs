using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoCAD_CSharp_plug_in_acCustomUI
{
    public partial class MyTopWindow : Form
    {
        public MyTopWindow()
        {
            InitializeComponent();
        }

        internal void showSelectedObjectsInfo(string info)
        {
            this.textBox1.Text = info;
        }
    }
}
