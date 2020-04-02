using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AcCommandTest
{
    public class FontUtils
    {
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

        public static void PutFontFiles()
        {
            string mcFontDir = Path.Combine(AssemblyDirectory, "Fonts");
            //TODO: 查找AutoCad安装目录
            string acFontDir = @"C:\Program Files\Autodesk\AutoCAD 2014\Fonts\";
            foreach (string file in Directory.EnumerateFiles(mcFontDir, "*.shx"))
            {
                string pure_file_name = Path.GetFileName(file);
                string dest_file = acFontDir + pure_file_name;
                if (!File.Exists(dest_file))
                {
                    File.Copy(file, dest_file);
                }
            }
        }

        public static void ProcessFont(Document doc)
        {
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId mcTextStyleId = GetMCTextStyle(db, tr);

                // 设置当前文字样式
                db.Textstyle = mcTextStyleId;

                // 替换模型空间的实体的文本样式
                BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                ChangeTextStyle(modelSpace, mcTextStyleId);

                // 替换图纸空间的实体的文本样式
                BlockTableRecord paperSpace = tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                ChangeTextStyle(paperSpace, mcTextStyleId);

                // 替换块定义的文本样式
                foreach (ObjectId id in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (!(btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference || btr.IsFromOverlayReference))
                    {
                        ChangeTextStyle(btr, mcTextStyleId);
                        // 刷新块引用
                        foreach (ObjectId brId in btr.GetBlockReferenceIds(true, true))
                        {
                            var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                            br.ResetBlock();
                        }
                        if (btr.IsDynamicBlock)
                        {
                            foreach (ObjectId btrId in btr.GetAnonymousBlockIds())
                            {
                                var anonymousBtr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                foreach (ObjectId brId in anonymousBtr.GetBlockReferenceIds(true, true))
                                {
                                    var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                                    br.ResetBlock();
                                }
                            }
                        }
                    }
                }

                tr.Commit();
            }
        }

        private static void ChangeTextStyle(BlockTableRecord btr, ObjectId textStyleId)
        {
            // 替换模型里所有文本的TextStyle
            foreach (ObjectId oid in btr)
            {
                if (oid.ObjectClass.DxfName == "TEXT")
                {
                    DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                    text.TextStyleId = textStyleId;
                }
                else if (oid.ObjectClass.DxfName == "MTEXT")
                {
                    MText text = (MText)oid.GetObject(OpenMode.ForWrite);
                    text.TextStyleId = textStyleId;
                }
                else if (oid.ObjectClass.DxfName == "DIMENSION")
                {
                    Dimension dim = (Dimension)oid.GetObject(OpenMode.ForWrite);
                    dim.TextStyleId = textStyleId;
                }
            }
        }

        /// <summary>
        /// 获取梦诚TextStyle的ObjectId，如没有则创建一个
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        private static ObjectId GetMCTextStyle(Database db, Transaction tr)
        {
            ObjectId mcTextStyleId = ObjectId.Null;
            TextStyleTable tst = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
            foreach (var oid in tst)
            {
                TextStyleTableRecord record = (TextStyleTableRecord)oid.GetObject(OpenMode.ForRead);
                if (record.Name == "@mc")
                {
                    mcTextStyleId = oid;
                    break;
                }
            }
            if (mcTextStyleId == ObjectId.Null)
            {
                TextStyleTableRecord newRecord = new TextStyleTableRecord();
                newRecord.Name = "@mc";
                newRecord.FileName = "@mc_symbol";
                newRecord.BigFontFileName = "@mc_bigfont";
                newRecord.TextSize = 0;
                tst.UpgradeOpen();
                tst.Add(newRecord);
                tr.AddNewlyCreatedDBObject(newRecord, true);
                mcTextStyleId = newRecord.Id;
            }
            return mcTextStyleId;
        }
    }
}
