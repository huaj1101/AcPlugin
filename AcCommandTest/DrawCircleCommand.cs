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

[assembly: CommandClass(typeof(AcCommandTest.DrawCircleCommand))]

namespace AcCommandTest
{
    public class DrawCircleCommand
    {
        public DrawCircleCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("DrawCircleCommand Init\r\n");
        }

        [CommandMethod("drawcircle")]
        public void DrawCircle()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //避免 eLockViolation 问题
                DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();

                //-------------------------------
                // 获取模型空间
                //-------------------------------
                BlockTable blockTbl = tr.GetObject(
                    db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(
                    blockTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                // 提示圆心
                PromptPointResult ppr = doc.Editor.GetPoint("请选取圆心: ");
                if (ppr.Status != PromptStatus.OK)
                {
                    m_DocumentLock.Dispose();
                    return;
                }
                Point3d ptCenter = ppr.Value;

                // 提示半径
                PromptDoubleResult pdr = doc.Editor.GetDouble("请输入半径: ");
                if (pdr.Status != PromptStatus.OK)
                {
                    m_DocumentLock.Dispose();
                    return;
                }
                double radius = pdr.Value;

                //-------------------------------
                // 创建圆
                //-------------------------------
                Circle circle = new Circle();
                circle.Center = ptCenter;
                circle.Normal = new Vector3d(0, 0, 1);
                circle.Radius = radius;

                //-------------------------------
                // 添加到模型空间并提交到数据库
                //-------------------------------
                modelSpace.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);
                tr.Commit();

                //避免 eLockViolation 问题
                m_DocumentLock.Dispose();
            }
        }
    }
}
