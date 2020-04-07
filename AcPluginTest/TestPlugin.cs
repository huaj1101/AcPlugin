using AcCommandTest;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

[assembly: ExtensionApplication(typeof(AcPluginTest.TestPlugin))]
[assembly: CommandClass(typeof(AcPluginTest.CommandLoader))]

namespace AcPluginTest
{
    public class TestPlugin : IExtensionApplication
    {
        public void Initialize()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("AcPluginTest插件初始化\n");
            CommandLoader loader = new CommandLoader();
            loader.LoadCommands();

            //处理字体
            try
            {
                FontUtils.PutFontFiles();
                FontUtils.ProcessFont(Application.DocumentManager.MdiActiveDocument);
                Application.DocumentManager.MdiActiveDocument.UserData.Add("mc_font", true);
            }
            catch (System.Exception e)
            {
                editor.WriteMessage(e.ToString());
            }
            //注册事件
            Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
        }

        void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (!e.Document.UserData.Contains("mc_font"))
            {
                try
                {
                    FontUtils.ProcessFont(e.Document);
                    e.Document.UserData.Add("mc_font", true);
                }
                catch (System.Exception ex)
                {
                    e.Document.Editor.WriteMessage(ex.ToString());
                }
            }
        }

        public void Terminate()
        {
            //do finalize
        }
    }

    public class CommandLoader
    {
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        [CommandMethod("mcload")]
        public void LoadCommands()
        {
            LoadAssemblyFile("AcCommandTest.dll");
            LoadAssemblyFile("EPPlus.dll");
        }

        private static void LoadAssemblyFile(string file)
        {
            byte[] buffer = System.IO.File.ReadAllBytes(Path.Combine(AssemblyDirectory, file));
            Assembly assembly = Assembly.Load(buffer);

            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("加载{0:s}\n", file);
        }

        /// <summary>
        /// 使用新菜单
        /// </summary>
        [CommandMethod("nm")]
        public void AssertNewMenu()
        {
            IAcadPreferences acadPreferences = (IAcadPreferences)Application.Preferences ;
            IAcadApplication acadApplication = (IAcadApplication)Application.AcadApplication;
            
            //acadApplication.Name;
            string profilesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                "Autodesk/AutoCAD 2014/R19.1/chs/Support/Profiles");
            const string newMenuName = "NEW_MENU_TEST__";
            string newProfile = Path.Combine(profilesDir, newMenuName);

            object profileNames;
            acadPreferences.Profiles.GetAllProfileNames(out profileNames);
            IList profileNameList = (IList)profileNames;
            if (!profileNameList.Contains(newMenuName))
            {
                acadPreferences.Profiles.CopyProfile(acadPreferences.Profiles.ActiveProfile, newMenuName);
            }

            if (!Directory.Exists(newProfile))
            {
                Directory.CreateDirectory(newProfile);
            }

            File.Copy(Path.Combine(AssemblyDirectory, "Profile.aws"), Path.Combine(newProfile, "Profile.aws"),true);
            acadPreferences.Profiles.ActiveProfile = newMenuName;
            acadApplication.Quit();
        }

        /// <summary>
        /// 重置为老菜单
        /// </summary>
        [CommandMethod("om")]
        public void RevertOldMenu()
        {
            IAcadPreferences acadPreferences = (IAcadPreferences)Application.Preferences;
            object profileNames;
            acadPreferences.Profiles.GetAllProfileNames(out profileNames);
            IList profileNameList = (IList)profileNames;
            acadPreferences.Profiles.ResetProfile(profileNameList[0] as string);
        }
    }

}
