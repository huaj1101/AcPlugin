using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AcStarter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private const string PROFILE_MC2020 = "MC2020";

        private static void TryRegistMyNetloadDll()
        {
            String mydir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            String dpath = Path.Combine(mydir, "AcCustomUI.dll");

            RegistryKey rootKey = Registry.CurrentUser;
            RegistryKey apps = rootKey.OpenSubKey(@"Software\Autodesk\AutoCAD\R19.1\ACAD-D001:804\Applications", true);
            if (apps != null)
            {
                RegistryKey MyPrograrm = apps.CreateSubKey(PROFILE_MC2020);
                MyPrograrm.SetValue("DESCRIPTION", "MC2020 应用", RegistryValueKind.String); //描述
                MyPrograrm.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord); //cad启动后自动加载
                MyPrograrm.SetValue("LOADER", dpath, RegistryValueKind.String); //dll的全路径
                MyPrograrm.SetValue("MANAGED", 1, RegistryValueKind.DWord);
            }
        }

        private static bool GetProfileImported()
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey profilesKey = key.OpenSubKey(@"Software\Autodesk\AutoCAD\R19.1\ACAD-D001:804\Profiles", false);
            Object val = profilesKey?.GetValue("");
            if (val != null && PROFILE_MC2020.Equals(val.ToString(), StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            else { 
                string[] subNames = profilesKey.GetSubKeyNames();
                return subNames.Contains(PROFILE_MC2020);
            }
        }

        private static bool GetAcadLocationFromRegistry(out string acadExe, out string exeArgs)
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey profilesKey = key.OpenSubKey(@"Software\Autodesk\AutoCAD\R19.1\ACAD-D001\Install", false);
            String mydir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var installDir = profilesKey?.GetValue("INSTALLDIR", null);
            if (installDir != null)
            {
                acadExe = Path.Combine(installDir.ToString(), "acad.exe");
                exeArgs = @"/nologo ";
                if (GetProfileImported())
                {
                    exeArgs += "/p \"" + PROFILE_MC2020 + "\"";
                }
                else
                {
                    String profileRegFile = Path.Combine(mydir, PROFILE_MC2020 + ".arg");

                    exeArgs += "/p \"" + profileRegFile + "\"";
                }

                exeArgs += " /s \"" + mydir + "\"";

                return true;
            }
            else
            {
                acadExe = null;
                exeArgs = null;
                return false;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //找到 AutoCAD 2014 安装位置
            string acadExe =null;
            string exeArgs = null;
            
            if (GetAcadLocationFromRegistry(out acadExe, out exeArgs))
            {
                Task task1 = new Task(() => StartAutoCAD(acadExe, exeArgs));
                task1.Start();
            }
            else
            {
                MessageBox.Show("AutoCAD 2014 未安装。");
                Application.Exit();
            }
        }

        private static void StartAutoCAD(string exeFile, string args)
        {
            System.Diagnostics.Process.Start(exeFile, args);

            Thread.Sleep(2000); //splash 停2秒
            Application.Exit();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            label1.Parent = pictureBox1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Parent = pictureBox1;
            TryRegistMyNetloadDll();
        }
    }
}
