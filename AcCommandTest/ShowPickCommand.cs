using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.ShowPickCommand))]

namespace AcCommandTest
{
    public class ShowPickCommand
    {
        public ShowPickCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("ShowPickCommand Init\r\n");
        }

        [CommandMethod("sp", CommandFlags.UsePickSet)]
        public void ShowPick()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 获取PickFirst选择集
                    PromptSelectionResult psr = editor.SelectImplied();

                    // 如果提示状态OK，说明启动命令前选择了对象;
                    if (psr.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("没有选中对象\n");
                        return;
                    }
                    SelectionSet sset = psr.Value;
                    editor.WriteMessage("选中{0:d}对象，第一个类型为{1:s}\n", sset.Count, sset[0].ObjectId.ObjectClass.DxfName);
                    StringBuilder sb = new StringBuilder();
                    foreach (SelectedObject so in sset)
                    {
                        if (so.ObjectId.ObjectClass.DxfName == "TEXT")
                        {
                            DBText text = (DBText)so.ObjectId.GetObject(OpenMode.ForRead);
                            sb.AppendLine(text.TextString);
                            string s = string.Format("( {0:f1}, {1:f1} ) ( {2:f1}, {3:f1} )", 
                                text.GeometricExtents.MinPoint.X,
                                text.GeometricExtents.MinPoint.Y,
                                text.GeometricExtents.MaxPoint.X,
                                text.GeometricExtents.MaxPoint.Y
                                );
                            sb.AppendLine(s);
                            //sb.Append(text.Position.ToString());
                        }
                        else if (so.ObjectId.ObjectClass.DxfName == "MTEXT")
                        {
                            MText text = (MText)so.ObjectId.GetObject(OpenMode.ForRead);
                            sb.AppendLine(text.Contents);
                            sb.AppendLine(text.Text);
                            string s = string.Format("( {0:f1}, {1:f1} ) ( {2:f1}, {3:f1} )",
                                text.GeometricExtents.MinPoint.X,
                                text.GeometricExtents.MinPoint.Y,
                                text.GeometricExtents.MaxPoint.X,
                                text.GeometricExtents.MaxPoint.Y
                                );
                            sb.AppendLine(s);
                            //sb.Append(text.Location.ToString());
                        }
                        else if (so.ObjectId.ObjectClass.DxfName == "LINE")
                        {
                            Line line = (Line)so.ObjectId.GetObject(OpenMode.ForRead);
                            sb.Append(line.StartPoint.ToString());
                            sb.Append(line.EndPoint.ToString());
                        }
                        else if (so.ObjectId.ObjectClass.DxfName == "LWPOLYLINE")
                        {
                            Polyline pLine = (Polyline)so.ObjectId.GetObject(OpenMode.ForRead);
                            for (int i = 0; i < pLine.NumberOfVertices; i++)
                            {
                                sb.Append(pLine.GetPoint2dAt(i).ToString());
                            }
                        }
                        else if (so.ObjectId.ObjectClass.DxfName == "INSERT")
                        {
                            Entity entity = (Entity)so.ObjectId.GetObject(OpenMode.ForWrite);
                            if (entity is BlockReference)
                            {
                                BlockReference br = entity as BlockReference;
                                sb.Append(br.Name);
                            }
                        }
                    }
                    editor.WriteMessage(sb.ToString());
                    trans.Commit();
                }
            }
            catch (System.Exception e)
            {
                editor.WriteMessage(e.ToString());
            }
        }
    }
}
