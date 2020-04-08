using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

                    //ShowConsoleTable(editor, table);
                    SaveToExcel(table);

                    trans.Commit();
                }
            }
            catch (System.Exception e)
            {
                editor.WriteMessage(e.StackTrace);
            }
        }

        private void SaveToExcel(Table table)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("sheet1");
                for (int i = 0; i < table.RowCount; i++)
                {
                    for (int j = 0; j < table.ColCount; j++)
                    {
                        if (table.Cells[i][j].Value != "")
                        {
                            worksheet.Cells[i + 1, j + 1].Value = table.Cells[i][j].Value;
                        }
                        if (table.Cells[i][j].CellType == TableCellType.MergedMaster)
                        {
                            worksheet.Cells[i + 1, j + 1, i + table.Cells[i][j].RowSpan, j + table.Cells[i][j].ColSpan].Merge = true;
                        }
                    }
                }
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Font.Name = "仿宋";
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].Style.WrapText = true;
                worksheet.Cells[1, 1, table.RowCount, table.ColCount].AutoFitColumns();

                var dialog = new System.Windows.Forms.SaveFileDialog();
                dialog.Filter = "Excel文件(*.xlsx)|*.xlsx";
                dialog.Title = "保存表格到Excel文件";
                dialog.FileName = "表格识别结果";
                
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    package.File = new FileInfo(dialog.FileName);
                    package.Save();
                    Process.Start(dialog.FileName);
                }
            }
        }

        private static void ShowConsoleTable(Editor editor, Table table)
        {
            ConsoleTableOptions options = new ConsoleTableOptions();
            ConsoleTable consoleTable = new ConsoleTable(options);
            consoleTable.AddColumn(table.Cells[0].Select((cell) =>
            {
                switch (cell.CellType)
                {
                    case TableCellType.Normal:
                        return "N";
                    case TableCellType.MergedMaster:
                        return string.Format("M {0:d} {1:d}", cell.RowSpan, cell.ColSpan);
                    case TableCellType.MergedSlave:
                        return "S";
                    default:
                        return "";
                }
            }).ToArray());
            for (int i = 1; i < table.RowCount; i++)
            {
                consoleTable.AddRow(table.Cells[i].Select((cell) =>
                {
                    switch (cell.CellType)
                    {
                        case TableCellType.Normal:
                            return "N";
                        case TableCellType.MergedMaster:
                            return string.Format("M {0:d} {1:d}", cell.RowSpan, cell.ColSpan);
                        case TableCellType.MergedSlave:
                            return "S";
                        default:
                            return "";
                    }
                }).ToArray());
            }
            editor.WriteMessage("\n" + consoleTable.ToMarkDownString());

            consoleTable = new ConsoleTable(options);
            consoleTable.AddColumn(table.Cells[0].Select((cell) => cell.Value));
            for (int i = 1; i < table.RowCount; i++)
            {
                consoleTable.AddRow(table.Cells[i].Select((cell) => cell.Value).ToArray());
            }
            editor.WriteMessage("\n" + consoleTable.ToMarkDownString());
        }
    }
}
