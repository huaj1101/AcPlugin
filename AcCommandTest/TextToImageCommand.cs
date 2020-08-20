using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcCommandTest.TextToImageCommand))]


namespace AcCommandTest
{
    class TextToImageCommand
    {
        [CommandMethod("ct", CommandFlags.UsePickSet)]
        public void CaptureText()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            List<TextCapture> list = new List<TextCapture>();

            using (DocumentLock m_DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 获取PickFirst选择集
                    PromptSelectionResult psr = editor.SelectImplied();

                    // 如果提示状态OK，说明启动命令前选择了对象;
                    if (psr.Status != PromptStatus.OK)
                    {
                        WriteMessage($"{psr.Status}没有选中对象\n");
                        return;
                    }
                    SelectionSet sset = psr.Value;
                    foreach (var oid in sset.GetObjectIds())
                    {
                        if (oid.ObjectClass.DxfName == "TEXT")
                        {
                            DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                            try
                            {
                                Bitmap img = EntityToImage(doc, tr, text);
                                TextCapture tc = new TextCapture(text.TextString, img);
                                list.Add(tc);
                            }
                            catch (System.Exception)
                            {
                                WriteMessage(text.TextString);
                                break;
                            }
                        }
                        else if (oid.ObjectClass.DxfName == "MTEXT")
                        {
                            MText text = (MText)oid.GetObject(OpenMode.ForWrite);
                            // WriteMessage(text.Text + "\n");
                            Bitmap img = EntityToImage(doc, tr, text);
                            TextCapture tc = new TextCapture(text.Text, img);
                            list.Add(tc);
                        }
                    }

                    tr.Commit();
                }
            }
            SaveImages(list);
        }

        [CommandMethod("cts")]
        public void CaptureTexts()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            List<TextCapture> list = new List<TextCapture>();

            using (DocumentLock m_DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 遍历模型空间
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    list.AddRange(EnumBlockTableRecord(doc, tr, modelSpace));

                    // 遍历块定义
                    foreach (ObjectId id in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (!(btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference || btr.IsFromOverlayReference))
                        {
                            list.AddRange(EnumBlockTableRecord(doc, tr, btr));
                        }
                    }
                    tr.Commit();
                }
            }
            SaveImages(list);
        }

        private void SaveImages(List<TextCapture> list)
        {
            string basePath = @"d:\cad_ocr_train_data\";
            string fileWithPath = Application.DocumentManager.MdiActiveDocument.Name;
            string pureName = Path.GetFileNameWithoutExtension(fileWithPath);
            string dir = Path.Combine(basePath, pureName);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            int index = 1;
            using (StreamWriter sw = new StreamWriter(Path.Combine(basePath, $"{pureName}.txt")))
            {

                for (int i = 0; i < list.Count; i++)
                {
                    if (list[i].Image != null && !string.IsNullOrEmpty(list[i].Text) && list[i].Text.Length > 1)
                    {
                        // WriteMessage($"{list[i].Text} {list[i].Text.Length}\n");
                        string img_file = string.Format("{0:s}-{1:d4}.png", pureName, index++);
                        string line = $"{img_file}\t{list[i].Text}";
                        list[i].Image.Save(Path.Combine(dir, img_file));
                        sw.WriteLine(line);
                    }
                }
            }
            WriteMessage($"{index - 1} images saved in {dir}");
        }

        private List<TextCapture> EnumBlockTableRecord(Document doc, Transaction tr, BlockTableRecord btr)
        {
            List<TextCapture> list = new List<TextCapture>();
            // 替换模型里所有文本的TextStyle
            foreach (ObjectId oid in btr)
            {
                if (oid.ObjectClass.DxfName == "TEXT")
                {
                    DBText text = (DBText)oid.GetObject(OpenMode.ForWrite);
                    if (string.IsNullOrEmpty(text.TextString))
                    {
                        continue;
                    }
                    try
                    {
                        Bitmap img = EntityToImage(doc, tr, text);
                        TextCapture tc = new TextCapture(text.TextString, img);
                        list.Add(tc);
                    }
                    catch (System.Exception e)
                    {
                        WriteMessage($"TEXT ERROR: {text.TextString} , 错误信息 {e}\n");
                        break;
                    }
                }
                else if (oid.ObjectClass.DxfName == "MTEXT")
                {
                    MText text = (MText)oid.GetObject(OpenMode.ForWrite);
                    if (string.IsNullOrEmpty(text.Text))
                    {
                        continue;
                    }
                    try
                    {
                        Bitmap img = EntityToImage(doc, tr, text);
                        TextCapture tc = new TextCapture(text.Text, img);
                        list.Add(tc);
                    }
                    catch (System.Exception e)
                    {
                        WriteMessage($"MTEXT ERROR: {text.Text} , 错误信息 {e}\n");
                        break;
                    }
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
                            list.AddRange(EnumBlockTableRecord(doc, tr, br.BlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord));
                        }
                    }
                }
            }
            return list;
        }

        private Bitmap EntityToImage(Document doc, Transaction tr, Entity entity)
        {
            //实体的宽度和高度
            var ext = entity.Bounds ?? entity.GeometricExtents;
            var width = ext.MaxPoint.X - ext.MinPoint.X;
            var height = ext.MaxPoint.Y - ext.MinPoint.Y;
            //只接受横向文本
            if (width < height)
            {
                return null;
            }

            ShowSingleLayer(doc, tr, entity.Layer);
            var view = doc.Editor.GetCurrentView();
            double scal1 = width / view.Width;
            double scal2 = height / view.Height;
            if (scal1 > scal2)
            {
                var viewHeight = view.Height * scal1;  //缩放到宽度适配，高度按比例修改
                Point3d minP = new Point3d(ext.MinPoint.X, ext.MaxPoint.Y - viewHeight, 0);
                Point3d maxP = ext.MaxPoint;
                ZoomWin(doc.Editor, minP, maxP); //缩放视口，使对象完整显示
            }
            else
            {
                var viewWidth= view.Width * scal2;  //缩放到高度适配，宽度按比例修改
                Point3d minP = ext.MinPoint;
                Point3d maxP = new Point3d(ext.MinPoint.X + viewWidth, ext.MaxPoint.Y, 0);
                ZoomWin(doc.Editor, minP, maxP); //缩放视口，使对象完整显示
            }

            //等比例缩放为高度32
            double newHeight = 32;
            double newWidth = (int)(newHeight / height * width);

            double ratioText = newWidth / newHeight;
            double ratioView = view.Width / view.Height;
            double bmpLen = newWidth;
            if (ratioText < ratioView)
            {
                bmpLen = bmpLen * ratioView / ratioText;
            }
            uint len = (uint)Math.Ceiling(bmpLen);
            Bitmap bmp = doc.CapturePreviewImage(len, len);
            Rectangle rec = new Rectangle(0, 0, (int)Math.Ceiling(newWidth), (int)newHeight);
            try
            {
                return bmp.Clone(rec, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            }
            catch (System.Exception)
            {
                WriteMessage($"vew: {view.Width}, {view.Height}   newwh: {newWidth}, {newHeight}   bmp: {bmp.Width}, {bmp.Height}\n");
                return null;
            }
        }

        private void WriteMessage(string message)
        {
            if (Application.DocumentManager.MdiActiveDocument != null && Application.DocumentManager.MdiActiveDocument.Editor != null)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(message);
            }
        }

        private void ShowSingleLayer(Document doc, Transaction trans, string layerName)
        {
            // 打开图层表
            LayerTable acLyrTbl = trans.GetObject(doc.Database.LayerTableId, OpenMode.ForWrite) as LayerTable;
            // 遍历图层
            foreach (ObjectId acObjId in acLyrTbl)
            {
                // 以读模式打开图层表记录以读打开图层表记录
                LayerTableRecord acLyrTblRec = trans.GetObject(acObjId, OpenMode.ForWrite) as LayerTableRecord;
                acLyrTblRec.IsOff = !layerName.Equals(acLyrTblRec.Name);
            }
        }

        private ViewTableRecord ZoomWin(Editor ed, Point3d min, Point3d max)
        {
            Point2d min2d = new Point2d(min.X, min.Y);
            Point2d max2d = new Point2d(max.X, max.Y);
            ViewTableRecord view = new ViewTableRecord();
            view.CenterPoint = min2d + ((max2d - min2d) / 2.0);
            view.Height = max2d.Y - min2d.Y;
            view.Width = max2d.X - min2d.X;
            ed.SetCurrentView(view);
            return view;
        }

    }

    class TextCapture
    {
        public string Text { get; set; }
        public Bitmap Image { get; set; }

        public TextCapture(string text, Bitmap image)
        {
            Text = text;
            Image = image;
        }
    }
}
