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
    public partial class MyDockBarLeft : UserControl
    {
        public MyDockBarLeft()
        {
            InitializeComponent();
        }

        public static void showSelectedObjectsInfo(String info)
        {
            if (myWindow != null)
            {
                myWindow.showSelectedObjectsInfo(info);

            }
        }

        private static MyTopWindow myWindow=null;
        private void button1_Click(object sender, EventArgs e)
        {
            if (myWindow == null)
            {
                myWindow = new MyTopWindow();
            }
            Application.ShowModelessDialog(myWindow);
        }

        private static MyModalDialog myModelDialog = null;
        private void button2_Click(object sender, EventArgs e)
        {
            if (myModelDialog == null)
            {
                myModelDialog = new MyModalDialog();
            }
            Application.ShowModalDialog(myModelDialog);
        }

 
        private void button3_Click(object sender, EventArgs e)
        {
   
        }
    }
}
