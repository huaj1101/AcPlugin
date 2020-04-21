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

[assembly: CommandClass(typeof(AcRecognize.EnumEntitiesCommand))]

namespace AcRecognize
{
    public class EnumEntitiesCommand
    {
        public EnumEntitiesCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("EnumEntitiesCommand Init\r\n");
        }

        [CommandMethod("ee")]
        public void EnumEntities()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            StringBuilder sb = new StringBuilder();
            using (DocumentLock m_DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 遍历模型空间
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    sb.AppendLine(EnumBlockTableRecord(modelSpace));

                    // 遍历块定义
                    foreach (ObjectId id in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (!(btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference || btr.IsFromOverlayReference))
                        {
                            sb.AppendLine(EnumBlockTableRecord(btr));
                        }
                    }
                    tr.Commit();
                }
            }
            using (StreamWriter sw = new StreamWriter(@"d:\enum_entities.txt"))
            {
                sw.Write(sb.ToString());
            }
            doc.Editor.WriteMessage(@"遍历结束，结果存储在 d:\enum_entities.txt");
        }

        private string EnumBlockTableRecord(BlockTableRecord btr)
        {
           StringBuilder sb = new StringBuilder();
            //
            foreach (ObjectId oid in btr)
            {
                if (oid.ObjectClass.DxfName == "TEXT")
                {
                    DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                    sb.AppendLine(text.TextString);
                }
                else if (oid.ObjectClass.DxfName == "MTEXT")
                {
                    MText text = (MText)oid.GetObject(OpenMode.ForWrite);
                    sb.AppendLine(text.Text);
                }
                else if (oid.ObjectClass.DxfName == "DIMENSION")
                {
                    Dimension dim = (Dimension)oid.GetObject(OpenMode.ForRead);
                    sb.AppendLine(dim.DimensionText);
                }
                else if (oid.ObjectClass.DxfName == "INSERT")
                {
                    Entity entity = (Entity)oid.GetObject(OpenMode.ForRead);
                    if (entity is BlockReference)
                    {
                        BlockReference br = entity as BlockReference;
                        // 处理未命名块（命名块在其他地方处理）
                        if (br.Name.StartsWith("*"))
                        {
                            sb.AppendLine(EnumBlockTableRecord(br.BlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord));
                        }
                    }
                }
            }
            return sb.ToString();
    }
    }
}
