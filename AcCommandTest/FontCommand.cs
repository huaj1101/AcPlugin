using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.FontCommand))]

namespace AcCommandTest
{
    public class FontCommand
    {
        public FontCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
        }

        [CommandMethod("ft")]
        public void ChangeFont()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            FontUtils.ProcessFont(doc);
        }

    }
}
