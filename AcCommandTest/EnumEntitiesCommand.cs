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

[assembly: CommandClass(typeof(AcCommandTest.EnumEntitiesCommand))]

namespace AcCommandTest
{
    public class EnumEntitiesCommand
    {
        public EnumEntitiesCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("EnumEntitiesCommand Init\r\n");
        }


        [CommandMethod("ee")]
        public void ChangeFont()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                StringBuilder sb = new StringBuilder();
                // 获取模型空间
                BlockTable blockTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                foreach (var item in blockTbl)
                {
                    
                }

                int index = 1;
                // 替换模型里所有文本的TextStyle
                foreach (ObjectId oid in modelSpace)
                {
                    DBObject obj = oid.GetObject(OpenMode.ForRead);
                    if (obj is Entity)
                    {
                        string line = string.Format("{0:d} {1:s} {2:s}", index++, oid.ObjectClass.DxfName, obj.GetType().Name);
                        sb.AppendLine(line);
                    }
                    if (obj is BlockReference)
                    {
                        BlockReference bref = obj as BlockReference;
                        DBObject block = bref.BlockTableRecord.GetObject(OpenMode.ForRead);
                        string line = string.Format("{0:d} {1:s} {2:s}", -1, bref.BlockTableRecord.ObjectClass.DxfName, block.GetType().Name);
                        sb.AppendLine(line);
                    }
                }
                //using (StreamWriter sw = new StreamWriter(@"d:\output.txt"))
                //{
                //    sw.Write(sb.ToString());
                //}
                doc.Editor.WriteMessage(sb.ToString());
                tr.Commit();

            }
        }

        [CommandMethod("eb")]
        public void ListBlockDefinitions()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId id in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (!(btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference || btr.IsFromOverlayReference))
                    {
                        ed.WriteMessage("\n" + btr.Name);
                        if (btr.HasAttributeDefinitions)
                        {
                            foreach (ObjectId attId in btr)
                            {
                                if (attId.ObjectClass.Name == "AcDbAttributeDefinition")
                                {
                                    var attDef = (AttributeDefinition)tr.GetObject(attId, OpenMode.ForRead);
                                    ed.WriteMessage("\n\t" + attDef.Tag);
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }
        }

    }
}
