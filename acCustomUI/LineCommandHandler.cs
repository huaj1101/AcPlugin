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
    public class LineCommandHandler : ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        [CommandMethod("ZsyNewLine")]
        public static void ZsyNewLine()
        {
            // 获取当前文档和数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // 启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //避免 eLockViolation 问题
                DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();

                // 以读模式打开 Block 表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 以写模式打开 Block 表记录 Model 空间
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;
                // Create a line
                Line acLine = new Line(new Point3d(50, 50, 0),
                    new Point3d(70, 70, 0));
                // 将新对象添加到块表记录和事务
                acBlkTblRec.AppendEntity(acLine);
                acTrans.AddNewlyCreatedDBObject(acLine, true);
                // 将新对象保存到数据库
                acTrans.Commit();

                //避免 eLockViolation 问题
                m_DocumentLock.Dispose();
            }
        }

        public void Execute(object parameter)
        {
            RibbonButton button = parameter as RibbonButton;
            if (button != null)
            {
                ZsyNewLine();
            }
        }
    }
}
