using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.ChangeFontCommand))]

namespace AcCommandTest
{
    public class ChangeFontCommand
    {
        public ChangeFontCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("ChangeFontCommand Init\r\n");
        }

        [CommandMethod("ft")]
        public void ChangeFont()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;
            //处理字体
            try
            {
                FontUtils.PutFontFiles();
                FontUtils.ProcessFont(Application.DocumentManager.MdiActiveDocument);
            }
            catch (System.Exception e)
            {
                editor.WriteMessage(e.ToString());
            }

        }
    }
}
