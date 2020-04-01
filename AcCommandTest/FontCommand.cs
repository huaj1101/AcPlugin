using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
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
            editor.WriteMessage("FontCommand Init\r\n");
        }

        [CommandMethod("ft")]
        public void ChangeFont()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //-------------------------------
                // 获取模型空间
                //-------------------------------
                BlockTable blockTbl = tr.GetObject(
                    db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(
                    blockTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                ObjectId mcTextStyleId = ObjectId.Null;
                TextStyleTable tst = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                foreach (var oid in tst)
                {
                    TextStyleTableRecord record = (TextStyleTableRecord)oid.GetObject(OpenMode.ForRead);
                    if (record.Name == "mc")
                    {
                        mcTextStyleId = oid;
                        break;
                    }
                }
                if (mcTextStyleId == ObjectId.Null)
                {
                    TextStyleTableRecord newRecord = new TextStyleTableRecord();
                    newRecord.Name = "mc";
                    newRecord.FileName = "tssdeng";
                    newRecord.BigFontFileName = "@~!HZTXT.SHX";
                    newRecord.TextSize = 0;
                    tst.UpgradeOpen();
                    tst.Add(newRecord);
                    tr.AddNewlyCreatedDBObject(newRecord, true);
                    mcTextStyleId = newRecord.Id;
                }

                StringBuilder sb = new StringBuilder();
                int i = 0;
                foreach (ObjectId oid in modelSpace)
                {
                    if (oid.ObjectClass.DxfName == "TEXT")
                    {
                        i++;
                        DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                        text.TextStyleId = mcTextStyleId;
                        // sb.AppendLine(string.Format("text{0:d}: {1:s}", i, text.TextString));
                    }
                }

                //doc.Editor.WriteMessage(sb.ToString());

                tr.Commit();

            }
        }
    }
}
