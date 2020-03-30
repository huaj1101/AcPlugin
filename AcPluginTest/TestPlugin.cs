using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
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
            editor.WriteMessage("AcPluginTest插件初始化");
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
    }
}
