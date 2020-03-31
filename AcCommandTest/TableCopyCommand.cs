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
                    AcTable table = AcTableParser.ParseTable(sset);

                    ConsoleTableOptions options = new ConsoleTableOptions();
                    ConsoleTable consoleTable = new ConsoleTable(options);
                    if (table.HasHeader)
                    {
                        consoleTable.AddColumn(table.Header);
                    }
                    for (int i = 0; i < table.DataRowCount; i++)
                    {
                        consoleTable.AddRow(table.Data[i]);
                    }
                    editor.WriteMessage("\n" + consoleTable.ToString());
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
