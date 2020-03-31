using System;
using System.Collections.Generic;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.Ribbon;
using Autodesk.Windows;
using System.Windows.Input;

namespace AutoCAD_CSharp_plug_in_acCustomUI
{
    public class CircleCommandHandler : ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        [CommandMethod("ZsyNewCircle")]
        public static void ZsyNewCircle()
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

                //-------------------------------
                // 创建圆
                //-------------------------------
                Circle circle = new Circle();
                circle.Center = new Point3d(100, 100, 0);
                circle.Normal = new Vector3d(0, 0, 1);
                circle.Radius = 20;

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

        public void Execute(object parameter)
        {
            RibbonButton button = parameter as RibbonButton;
            if (button != null)
            {
                ZsyNewCircle();
            }
        }
    }
}
