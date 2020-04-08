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
                    Table table = AcTableParser.ParseTable(sset.GetObjectIds());

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
                editor.WriteMessage(e.ToString());
            }
        }

        private void SaveToExcel(Table table)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("sheet1");
                for (int i = 0; i < table.RowCount; i++)
                {
                    for (int j = 0; j < table.ColCount; j++)
                    {
                        if (table.Cells[i][j].Value != "")
                        {
                            ws.Cells[i + 1, j + 1].Value = table.Cells[i][j].Value;
                        }
                        if (table.Cells[i][j].CellType == TableCellType.MergedMaster)
                        {
                            ws.Cells[i + 1, j + 1, i + table.Cells[i][j].RowSpan, j + table.Cells[i][j].ColSpan].Merge = true;
                        }
                    }
                }

                ws.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, table.RowCount, table.ColCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; 
                ws.Cells.Style.Font.Name = "仿宋";
                ws.Cells.Style.Font.Size = 10;
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                ws.Cells.AutoFitColumns();
                ws.Cells.Style.WrapText = true;
                for (int i = 1; i <= table.ColCount; i++)
                {
                    ws.Column(i).Width = ws.Column(i).Width * 0.88;
                }

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
