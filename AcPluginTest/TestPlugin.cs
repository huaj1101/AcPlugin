using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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

    }

}
