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
            MyTopWindow.showSelectedObjectsInfo(info);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyTopWindow.showModeless();
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
            LineCommandHandler.NewLine();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CircleCommandHandler.NewCircle();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new UiManager().McUiSwitch();
        }
    }
}
