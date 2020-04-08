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

        internal static void showSelectedObjectsInfo(string info)
        {
            if (myWindow == null)
            {
                myWindow = new MyTopWindow();
            }
            myWindow.textBox1.Text = "选定对象ID列表：\r\n" + info;
        }

        private static MyTopWindow myWindow = null;
        internal static void showModeless()
        {
            if (myWindow == null)
            {
                myWindow = new MyTopWindow();
            }
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(myWindow);
        }

        private void MyTopWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }
    }
}
