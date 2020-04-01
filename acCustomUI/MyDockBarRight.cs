using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_CSharp_plug_in_acCustomUI
{
    public partial class MyDockBarRight : UserControl
    {
        public MyDockBarRight()
        {
            InitializeComponent();
            this.elementHost1.Child = new MyWpfUserControl(); //嵌入 WPF 用户控件
        }
    }
}
