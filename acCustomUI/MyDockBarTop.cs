﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoCAD_CSharp_plug_in_acCustomUI
{
    public partial class MyDockBarTop : UserControl
    {
        public MyDockBarTop()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new UiManager().McUiSwitch();
        }
    }
}
