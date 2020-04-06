using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.TableCopyCommand))]

namespace AcCommandTest
{
    public class TableCopyCommand
    {
        public TableCopyCommand()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("TableCopyCommand Init\r\n");
        }

        [CommandMethod("tc", CommandFlags.UsePickSet)]
        public void TableCopy()
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
                    Table table = null;
                    // 块类型的表格
                    if (sset.Count == 1 && sset[0].ObjectId.ObjectClass.DxfName == "INSERT")
                    {
                        Entity entity = sset[0].ObjectId.GetObject(OpenMode.ForRead) as Entity;
                        if (entity is BlockReference)
                        {
                            BlockReference br = entity as BlockReference;
                            BlockTableRecord btr = br.BlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                            table = AcTableParser.ParseTable(btr);
                        }
                    }
                    else
                    {
                        table = AcTableParser.ParseTable(sset.GetObjectIds());
                    }

                    if (table == null || table.RowCount == 0)
                    {
                        editor.WriteMessage("\n无法识别表格");
                    }
                    ConsoleTableOptions options = new ConsoleTableOptions();
                    ConsoleTable consoleTable = new ConsoleTable(options);
                    consoleTable.AddColumn(table.Cells[0].Select((cell) => cell.Value));
                    //consoleTable.AddColumn(table.Cells[0].Select((cell) =>
                    //{
                    //    switch (cell.CellType)
                    //    {
                    //        case TableCellType.Normal:
                    //            return "N";
                    //        case TableCellType.MergedMaster:
                    //            return string.Format("M {0:d} {1:d}", cell.RowSpan, cell.ColSpan);
                    //        case TableCellType.MergedSlave:
                    //            return "S";
                    //        default:
                    //            return "";
                    //    }
                    //}).ToArray());
                    for (int i = 1; i < table.RowCount; i++)
                    {
                        consoleTable.AddRow(table.Cells[i].Select((cell) => cell.Value).ToArray());
                        //consoleTable.AddRow(table.Cells[i].Select((cell) => 
                        //    {
                        //        switch (cell.CellType)
                        //        {
                        //            case TableCellType.Normal:
                        //                return "N";
                        //            case TableCellType.MergedMaster:
                        //                return string.Format("M {0:d} {1:d}", cell.RowSpan, cell.ColSpan);
                        //            case TableCellType.MergedSlave:
                        //                return "S";
                        //            default:
                        //                return "";
                        //        }
                                
                        //    }).ToArray());
                    }
                    editor.WriteMessage("\n" + consoleTable.ToMarkDownString());
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
