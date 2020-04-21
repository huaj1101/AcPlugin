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
        private const string REG_VERSION_SUB_KEY_HKLM = @"R23.1\ACAD-3001";
        private const string REG_VERSION_SUB_KEY_ZHCN = @"R23.1\ACAD-3001:804";
        private const string PROFILE_MC2020 = "MCSL2020";

        private static bool zh_cn_lang_installed = false;
        private static string cadInstallDir = null;
        private static string acadExe = null;
        private static string exeArgs = null;
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 本处理应由产品安装包来处理
        /// </summary>
        private static void TryRegistMyNetloadDll__()
        {
            var mydir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var myDll = Path.Combine(mydir, "AcCustomUI.dll");
            
            var dpath = Path.Combine(cadInstallDir, @"Plugins\mcsl2020\AcCustomUI.dll");
            if (!File.Exists(dpath)) //复制 dll 到 AutoCAD 的 Plugins 目录下，则AutoCAD 启动后加载 dll 不再有确认提示
            {
                var dDir = Path.Combine(cadInstallDir, @"Plugins\mcsl2020");
                if (!Directory.Exists(dDir))
                {
                    Directory.CreateDirectory(dDir);
                }
                File.Copy(myDll, dpath);
            }

            RegistryKey rootKey = Registry.CurrentUser;
            RegistryKey apps = rootKey.OpenSubKey(@"Software\Autodesk\AutoCAD\"+ REG_VERSION_SUB_KEY_ZHCN +@"\Applications", true);
            if (apps == null) return;
            RegistryKey MyPrograrm = apps.CreateSubKey(PROFILE_MC2020);
            MyPrograrm.SetValue("DESCRIPTION", "MCSL2020 应用", RegistryValueKind.String); //描述
            MyPrograrm.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord); //cad启动后自动加载
            MyPrograrm.SetValue("LOADER", dpath, RegistryValueKind.String); //dll的全路径
            MyPrograrm.SetValue("MANAGED", 1, RegistryValueKind.DWord);
        }


        private static void TryRegistMyNetloadDll()
        {
            var mydir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var dpath = Path.Combine(mydir, "AcCustomUI.dll");
            RegistryKey rootKey = Registry.CurrentUser;
            RegistryKey apps = rootKey.OpenSubKey(@"Software\Autodesk\AutoCAD\" + REG_VERSION_SUB_KEY_ZHCN + @"\Applications", true);
            if (apps == null) return;
            RegistryKey MyPrograrm = apps.CreateSubKey(PROFILE_MC2020);
            MyPrograrm.SetValue("DESCRIPTION", "MCSL2020 应用", RegistryValueKind.String); //描述
            MyPrograrm.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord); //cad启动后自动加载
            MyPrograrm.SetValue("LOADER", dpath, RegistryValueKind.String); //dll的全路径
            MyPrograrm.SetValue("MANAGED", 1, RegistryValueKind.DWord);
        }

        private static bool GetProfileImported()
        {
            RegistryKey key = Registry.CurrentUser;
            RegistryKey profilesKey = key.OpenSubKey(@"Software\Autodesk\AutoCAD\"+ REG_VERSION_SUB_KEY_ZHCN +@"\Profiles", false);
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

        private static bool GetAcadLocationFromRegistry()
        {
            RegistryKey rootKeyHKCU = Registry.CurrentUser;
            RegistryKey key_zhcn = rootKeyHKCU.OpenSubKey(@"Software\Autodesk\AutoCAD\" + REG_VERSION_SUB_KEY_ZHCN, false);
            zh_cn_lang_installed = key_zhcn != null;
            if (!zh_cn_lang_installed)
            {
                return false;
            }

            
            RegistryKey key = Registry.LocalMachine;
            RegistryKey profilesKey = key.OpenSubKey(@"Software\Autodesk\AutoCAD\"+ REG_VERSION_SUB_KEY_HKLM + @"\Install", false);
            var mydir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var cadInstallDirObj = profilesKey?.GetValue("INSTALLDIR", null);
            if (cadInstallDirObj != null)
            {
                cadInstallDir = cadInstallDirObj.ToString();
                acadExe = Path.Combine(cadInstallDir, "acad.exe");
                exeArgs = "/nologo /product ACAD /language \"zh-CN\" ";
                if (GetProfileImported())
                {
                    exeArgs += "/p \"" + PROFILE_MC2020 + "\"";
                }
                else
                {
                    var profileRegFile = Path.Combine(mydir, PROFILE_MC2020 + ".arg");

                    exeArgs += "/p \"" + profileRegFile + "\"";
                }

                var supdir = Path.Combine(cadInstallDir.ToString(), "support");
                exeArgs += " /s \"" + mydir + ";" + supdir + "\"";

                return true;
            }
            else
            {
                return false;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //找到 AutoCAD 2020 安装位置
            if (acadExe!=null)
            {
                Task task1 = new Task(() => StartAutoCAD(acadExe, exeArgs));
                task1.Start();
            }
            else if(!zh_cn_lang_installed)
            {
                MessageBox.Show("未安装 AutoCAD 2020 中文语言包。");
                Application.Exit();
            }
            else
            {
                MessageBox.Show("AutoCAD 2020 未安装。");
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
            GetAcadLocationFromRegistry();
            TryRegistMyNetloadDll();
        }
    }
}
