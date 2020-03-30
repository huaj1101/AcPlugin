using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.HelloWorldCommand))]

namespace AcCommandTest
{
    public class HelloWorldCommand
    {
        public HelloWorldCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("HellowWorldCommand Init\r\n");
        }

        [CommandMethod("hellowworld")]
        public void HelloWorld()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("hellow world");
        }
    }
}
