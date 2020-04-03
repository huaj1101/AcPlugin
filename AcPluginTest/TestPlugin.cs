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
            editor.WriteMessage("AcPluginTest插件初始化\r\n");
            CommandLoader loader = new CommandLoader();
            loader.LoadCommands();

            //处理字体
            FontUtils.PutFontFiles();
            foreach (Document doc in Application.DocumentManager)
            {
                FontUtils.ProcessFont(doc);
            }
            //注册事件
            Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
        }


        void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            FontUtils.ProcessFont(e.Document);
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
            byte[] buffer = System.IO.File.ReadAllBytes(Path.Combine(AssemblyDirectory, "AcCommandTest.dll"));
            Assembly assembly = Assembly.Load(buffer);

            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("加载AcCommandTest.dll");
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
