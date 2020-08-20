using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace AcCommandTest
{
    public class FontUtils
    {
        private static readonly double WIDTH_SCALE = 0.71;
        private static readonly double HEIGHT_SCALE = 0.94;

        private static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        private static string AutoCadDirectory
        {
            get
            {
                string acFile = Process.GetCurrentProcess().MainModule.FileName;
                return Path.GetDirectoryName(acFile);
            }
        }

        public static void PutFontFiles()
        {
            string mcFontDir = Path.Combine(AssemblyDirectory, "Fonts");
            string acFontDir = Path.Combine(AutoCadDirectory, "Fonts");
            foreach (string file in Directory.EnumerateFiles(mcFontDir, "*.shx"))
            {
                string pure_file_name = Path.GetFileName(file);
                string dest_file = Path.Combine(acFontDir, pure_file_name);
                try
                {
                    if (!File.Exists(dest_file))
                    {
                        File.Copy(file, dest_file, true);
                    }
                }
                catch (System.IO.IOException e)
                {
                    //覆盖不掉提示一下
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            }
        }

        public static void ProcessFont(Document doc)
        {
            if (!doc.IsActive)
            {
                throw new Exception("只能处理活动状态的Document");
            }
            doc.Editor.WriteMessage("\n正在处理字体。\n");
            DateTime start = DateTime.Now;

            Database db = doc.Database;

            using (DocumentLock m_DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    TextStyleTable tst = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    foreach (var oid in tst)
                    {
                        TextStyleTableRecord record = (TextStyleTableRecord)oid.GetObject(OpenMode.ForRead);
                        if (record.FileName != "@mc_symbol")
                        {
                            doc.UserData[record.Name] = true;
                            record.UpgradeOpen();
                            record.FileName = "@mc_symbol";
                            record.BigFontFileName = "@mc_bigfont";
                            record.XScale = record.XScale * WIDTH_SCALE;
                        }
                        else
                        {
                            doc.UserData[record.Name] = false;
                        }
                    }

                    // 替换模型空间的实体的文本样式
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Redraw(doc, modelSpace);

                    // 替换块定义的文本样式
                    foreach (ObjectId id in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (!(btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference || btr.IsFromOverlayReference))
                        {
                            Redraw(doc, btr);
                            // 刷新块引用
                            foreach (ObjectId brId in btr.GetBlockReferenceIds(true, true))
                            {
                                var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                                br.RecordGraphicsModified(true);
                            }
                            if (btr.IsDynamicBlock)
                            {
                                foreach (ObjectId btrId in btr.GetAnonymousBlockIds())
                                {
                                    var anonymousBtr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                    foreach (ObjectId brId in anonymousBtr.GetBlockReferenceIds(true, true))
                                    {
                                        var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                                        br.RecordGraphicsModified(true);
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            doc.Editor.WriteMessage("\n处理字体成功，用时{0:f2}s。\n", (DateTime.Now - start).TotalSeconds);
        }

        private static void Redraw(Document doc, BlockTableRecord btr)
        {
            Regex regex = new Regex(@"\\[Ff][^;]+;");
            
            // 替换模型里所有文本的TextStyle
            foreach (ObjectId oid in btr)
            {
                if (oid.ObjectClass.DxfName == "TEXT")
                {
                    DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                    if (doc.UserData.ContainsKey(text.TextStyleName) && (bool)doc.UserData[text.TextStyleName])
                    {
                        text.WidthFactor = text.WidthFactor * WIDTH_SCALE;
                        text.Height = text.Height * HEIGHT_SCALE;
                    }
                }
                else if (oid.ObjectClass.DxfName == "MTEXT")
                {
                    MText text = (MText)oid.GetObject(OpenMode.ForWrite);
                    if (doc.UserData.ContainsKey(text.TextStyleName) && (bool)doc.UserData[text.TextStyleName])
                    {
                        text.TextHeight = text.TextHeight * HEIGHT_SCALE;
                        text.LineSpacingFactor = text.LineSpacingFactor * HEIGHT_SCALE;
                        // 多行文本中可以嵌入字体，去掉这个信息
                        text.Contents = regex.Replace(text.Contents, "");
                    }
                }
                else if (oid.ObjectClass.DxfName == "DIMENSION")
                {
                    Dimension dim = (Dimension)oid.GetObject(OpenMode.ForRead);
                    dim.Draw();
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
                            Redraw(doc, br.BlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord);
                        }
                        br.RecordGraphicsModified(true);
                    }
                }
            }
        }
    }
}
