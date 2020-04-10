using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcCommandTest
{
    /// <summary>
    /// 解析AucoCad里的表格，目前只支持由正交的直线、多段线画的表格，支持以下特性：
    /// 1. 识别有合并单元格的列头，最终解析成单列头竖线分隔
    /// 2. 识别合计行
    /// 3. 格子中有0-n个单行/多行文本
    /// </summary>
    public class AcTableParser
    {
        private List<AcText> _texts = new List<AcText>();
        private List<AcTableLine> _tableHLines = new List<AcTableLine>();
        private List<AcTableLine> _tableVLines = new List<AcTableLine>();
        private AcTableCell[,] _cells;
        private int _rowCount;
        private int _colCount;
        /// <summary>
        /// 容错大小
        /// 有时候画的格线有误差搭接不上，增加一些容错性
        /// 表格有大有小，所以容错大小是根据表格大小计算出来的
        /// </summary>
        private double _tolerance;

        /// <summary>
        /// 解析选中对象为一个表格，如果无法识别会抛出AcTableParseException异常，需调用方处理异常
        /// </summary>
        /// <param name="objects"></param>
        /// <returns></returns>
        public static Table ParseTable(IEnumerable objects)
        {
            return new AcTableParser().Parse(objects);
        }

        /// <summary>
        /// 解析表格
        /// </summary>
        /// <param name="objects"></param>
        /// <returns></returns>
        private Table Parse(IEnumerable objects)
        {
            ParseObjects(objects);
            SortObjects();
            CalcTolerance();
            //System.Windows.Forms.MessageBox.Show(_tableHLines.Count.ToString());
            //System.Windows.Forms.MessageBox.Show(_tableHLines.Last().ToString());
            //由于清理算法原因，次数特意清理多次
            CleanLines();
            CleanLines();
            CleanLines();
            //string s = "";
            //foreach (var item in _tableVLines)
            //{
            //    s += item.XorY.ToString() + ",";
            //}
            //System.Windows.Forms.MessageBox.Show(s);
            BuildCells();
            PutTextToCells();
            CalcCellMerge();
            CalcCellsValue();
            return GenerateResult();
        }

        /// <summary>
        /// 计算容错大小
        /// 具体计算办法是平均格子宽度、高度的1/10
        /// </summary>
        private void CalcTolerance()
        {
            double hMaxLen = _tableHLines.Max(line => line.Length);
            double hTolerance = hMaxLen / (_tableVLines.Count - 1) / 10;
            double vMaxLen = _tableVLines.Max(line => line.Length);
            double vTolerance = vMaxLen / (_tableHLines.Count - 1) / 10;
            _tolerance = Math.Min(hTolerance, vTolerance);
            //System.Windows.Forms.MessageBox.Show(_tolerance.ToString());
        }

        /// <summary>
        /// 清理无效的格线
        /// 有时候表格里会有非格线的直线或多段线（比如钢筋图样），需要清理掉
        /// 清理的原则是要求格线的头或尾，必须在某个正交方向的格线上
        /// </summary>
        private void CleanLines()
        {
            //清理横线
            AcTableLine preHLine = null;
            foreach (AcTableLine hLine in _tableHLines.ToArray())
            {
                hLine.CrossPointCount = 0;
                foreach (AcTableLine vLine in _tableVLines)
                {
                    Point2d pt = new Point2d(vLine.XorY, hLine.XorY);
                    if (hLine.HasSegmentOn(pt, _tolerance) && vLine.HasSegmentOn(pt, _tolerance))
                    {
                        hLine.CrossPointCount += 1;
                    }
                }
                if (hLine.CrossPointCount <= 1)
                {
                    _tableHLines.Remove(hLine);
                }
                else if (preHLine != null && preHLine.XorY - hLine.XorY <= _tolerance) //表格容错，挨的太近的格线，删除掉一根交叉点小的
                {
                    //string s = string.Format("{0:f3},{1:f3},{2:f3},{3:f3}", preHLine.XorY, hLine.XorY, preHLine.XorY - hLine.XorY, _tolerance);
                    //System.Windows.Forms.MessageBox.Show(s);
                    if (preHLine.CrossPointCount < hLine.CrossPointCount)
                    {
                        _tableHLines.Remove(preHLine);
                        preHLine = hLine;
                    }
                    else
                    {
                        _tableHLines.Remove(hLine);
                    }
                }
                else
                {
                    preHLine = hLine;
                }
            }
            //清理竖线
            AcTableLine preVLine = null;
            foreach (AcTableLine vLine in _tableVLines.ToArray())
            {
                vLine.CrossPointCount = 0;
                foreach (AcTableLine hLine in _tableHLines)
                {
                    Point2d pt = new Point2d(vLine.XorY, hLine.XorY);
                    if (hLine.HasSegmentOn(pt, _tolerance) && vLine.HasSegmentOn(pt, _tolerance))
                    {
                        vLine.CrossPointCount += 1;
                    }
                }
                if (vLine.CrossPointCount <= 1)
                {
                    _tableVLines.Remove(vLine);
                }
                else if (preVLine != null && vLine.XorY - preVLine.XorY < _tolerance) //表格容错，挨的太近的格线，删除掉一根交叉点小的
                {
                    if (preVLine.CrossPointCount < vLine.CrossPointCount)
                    {
                        _tableVLines.Remove(preVLine);
                        preVLine = vLine;
                    }
                    else
                    {
                        _tableVLines.Remove(vLine);
                    }
                }
                else
                {
                    preVLine = vLine;
                }
            }
            //有时候表格会有双层外框，清理外层
            if (_tableHLines.Count > 2 || _tableVLines.Count > 2)
            {
                if (_tableHLines[0].CrossPointCount <= 2 &&
                    _tableHLines.Last().CrossPointCount <= 2 &&
                    _tableVLines[0].CrossPointCount <= 2 &&
                    _tableVLines.Last().CrossPointCount <= 2)
                {
                    _tableHLines.RemoveAt(0);
                    _tableHLines.RemoveAt(_tableHLines.Count - 1);
                    _tableVLines.RemoveAt(0);
                    _tableVLines.RemoveAt(_tableVLines.Count - 1);
                }
            }

            //清理只有2个交点并且不等于表格宽和高的格线
            double tableWidth = _tableHLines.Max(line => line.Length);
            foreach (AcTableLine hLine in _tableHLines.ToArray())
            {
                if (hLine.CrossPointCount <= 2 && CommandUtils.Compare(hLine.Length, tableWidth) < 0)
                {
                    //System.Windows.Forms.MessageBox.Show(hLine.Length.ToString());
                    _tableHLines.Remove(hLine);
                }
            }
            double tableHeight = _tableVLines.Max(line => line.Length);
            foreach (AcTableLine vLine in _tableVLines.ToArray())
            {
                if (vLine.CrossPointCount <= 2 && CommandUtils.Compare(vLine.Length, tableHeight) < 0)
                {
                    _tableVLines.Remove(vLine);
                }
            }
        }

        /// <summary>
        /// 计算单元格合并
        /// </summary>
        private void CalcCellMerge()
        {
            for (int i = 0; i < _rowCount; i++)
            {
                for (int j = 0; j < _colCount; j++)
                {
                    if (_cells[i, j].InnerCell.CellType == TableCellType.MergedSlave)
                    {
                        if (_cells[i, j].InnerCell.Row == _cells[i, j].InnerCell.MasterCell.Row &&
                            !_cells[i, j].RightLine.HasSegmentOn(_cells[i, j].Center.Y, _tolerance) &&
                            j < _colCount - 1 &&
                            _cells[i, j + 1].InnerCell.CellType == TableCellType.Normal)
                        {
                            _cells[i, j].InnerCell.MasterCell.ColSpan += 1;
                            for (int row = i; row < i + _cells[i, j].InnerCell.MasterCell.RowSpan; row++)
                            {
                                _cells[row, j + 1].InnerCell.MasterCell = _cells[i, j].InnerCell.MasterCell;
                                _cells[row, j + 1].InnerCell.CellType = TableCellType.MergedSlave;
                            }
                        }
                        else if (_cells[i, j].InnerCell.Col == _cells[i, j].InnerCell.MasterCell.Col &&
                            !_cells[i, j].BottomLine.HasSegmentOn(_cells[i, j].Center.X, _tolerance) &&
                            i < _rowCount - 1 &&
                            _cells[i + 1, j].InnerCell.CellType == TableCellType.Normal)
                        {
                            _cells[i, j].InnerCell.MasterCell.RowSpan += 1;
                            for (int col = j; col < j + _cells[i, j].InnerCell.MasterCell.ColSpan; col++)
                            {
                                _cells[i + 1, col].InnerCell.MasterCell = _cells[i, j].InnerCell.MasterCell;
                                _cells[i + 1, col].InnerCell.CellType = TableCellType.MergedSlave;
                            }
                        }
                    }
                    else if (_cells[i, j].InnerCell.CellType == TableCellType.Normal)
                    {
                        if (!_cells[i, j].RightLine.HasSegmentOn(_cells[i, j].Center.Y, _tolerance) && j < _colCount - 1)
                        {
                            _cells[i, j].InnerCell.CellType = TableCellType.MergedMaster;
                            _cells[i, j].InnerCell.ColSpan = 2;
                            _cells[i, j + 1].InnerCell.MasterCell = _cells[i, j].InnerCell;
                            _cells[i, j + 1].InnerCell.CellType = TableCellType.MergedSlave;
                        }
                        if (!_cells[i, j].BottomLine.HasSegmentOn(_cells[i, j].Center.X, _tolerance) && i < _rowCount - 1)
                        {
                            _cells[i, j].InnerCell.CellType = TableCellType.MergedMaster;
                            _cells[i, j].InnerCell.RowSpan = 2;
                            _cells[i + 1, j].InnerCell.MasterCell = _cells[i, j].InnerCell;
                            _cells[i + 1, j].InnerCell.CellType = TableCellType.MergedSlave;
                            if (_cells[i, j].InnerCell.ColSpan == 2 && j < _colCount - 1)
                            {
                                _cells[i + 1, j + 1].InnerCell.MasterCell = _cells[i, j].InnerCell;
                                _cells[i + 1, j + 1].InnerCell.CellType = TableCellType.MergedSlave;
                            }
                        }
                    }
                    else
                    {
                        throw new AcTableParseException("表格解析错误");
                    }
                }
            }
            //处理合并后Text的归属
            foreach (AcTableCell cell in _cells)
            {
                if (cell.InnerCell.CellType == TableCellType.MergedSlave && cell.Texts.Count > 0)
                {
                    AcTableCell masterCell = _cells[cell.InnerCell.MasterCell.Row, cell.InnerCell.MasterCell.Col];
                    masterCell.Texts.AddRange(cell.Texts);
                    cell.Texts.Clear();
                }
            }

        }

        /// <summary>
        /// 生成AcTable类型的返回结果
        /// </summary>
        /// <returns></returns>
        private Table GenerateResult()
        {
            Table table = new Table();
            table.Cells = new TableCell[_rowCount][];
            for (int i = 0; i < _rowCount; i++)
            {
                table.Cells[i] = new TableCell[_colCount];
                for (int j = 0; j < _colCount; j++)
                {
                    table.Cells[i][j] = _cells[i, j].InnerCell;
                }
            }

            return table;
        }

        /// <summary>
        /// 计算每个格子的文本值
        /// </summary>
        private void CalcCellsValue()
        {
            foreach (AcTableCell cell in _cells)
            {
                cell.CalcCellValue();
            }
        }

        /// <summary>
        /// 将文本信息放入对应的格子中
        /// </summary>
        private void PutTextToCells()
        {
            int startCol = 0;
            foreach (AcText text in _texts)
            {
                bool put = false;
                for (int j = startCol; j < _colCount; j++)
                {
                    for (int i = 0; i < _rowCount; i++)
                    {
                        if (CommandUtils.PointInRect(text.Position, _cells[i, j].TopLeft, _cells[i, j].BottomRight))
                        {
                            _cells[i, j].Texts.Add(text);
                            put = true;
                            break;
                        }
                    }
                    if (put)
                    {
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 根据格线构建出格子
        /// </summary>
        private void BuildCells()
        {
            _rowCount = _tableHLines.Count - 1;
            _colCount = _tableVLines.Count - 1;
            _cells = new AcTableCell[_rowCount, _colCount];
            for (int i = 0; i < _tableHLines.Count - 1; i++)
            {
                for (int j = 0; j < _tableVLines.Count - 1; j++)
                {
                    AcTableCell tc = new AcTableCell();
                    tc.InnerCell.Row = i;
                    tc.InnerCell.Col = j;
                    tc.TopLeft = new Point2d(_tableVLines[j].XorY, _tableHLines[i].XorY);
                    tc.BottomRight = new Point2d(_tableVLines[j + 1].XorY, _tableHLines[i + 1].XorY);
                    tc.Center = new Point2d((tc.TopLeft.X + tc.BottomRight.X) / 2, (tc.TopLeft.Y + tc.BottomRight.Y) / 2);
                    tc.LeftLine = _tableVLines[j];
                    tc.RightLine = _tableVLines[j + 1];
                    tc.TopLine = _tableHLines[i];
                    tc.BottomLine = _tableHLines[i + 1];
                    _cells[i, j] = tc;
                }
            }
        }

        /// <summary>
        /// 识别选中的对象，将其变为内部结构
        /// </summary>
        /// <param name="objects"></param>
        private void ParseObjects(IEnumerable objects, double baseX = 0, double baseY = 0, double xScale = 1, double yScale = 1)
        {
            foreach (ObjectId oid in objects)
            {
                switch (oid.ObjectClass.DxfName)
                {
                    case "TEXT":
                        DBText text = (DBText)oid.GetObject(OpenMode.ForRead);
                        Point2d ptTextCenter = CommandUtils.GetCenterPoint(text);
                        AcText at1 = new AcText(text.TextString, new Point2d(baseX + xScale * ptTextCenter.X, baseY + yScale * ptTextCenter.Y), yScale * text.Height);
                        if (!TextExists(at1))
                        {
                            _texts.Add(at1);
                        }
                        break;
                    case "MTEXT":
                        MText mText = (MText)oid.GetObject(OpenMode.ForRead);
                        Point2d ptMTextCenter = CommandUtils.GetCenterPoint(mText);
                        AcText at2 = new AcText(mText.Text, new Point2d(baseX + xScale * ptMTextCenter.X, baseY + yScale * ptMTextCenter.Y), yScale * mText.Height);
                        if (!TextExists(at2))
                        {
                            _texts.Add(at2);
                        }
                        break;
                    case "LINE":
                        Line line = (Line)oid.GetObject(OpenMode.ForRead);
                        ParseLine(new Point2d(baseX + xScale * line.StartPoint.X, baseY + yScale * line.StartPoint.Y),
                            new Point2d(baseX + xScale * line.EndPoint.X, baseY + yScale * line.EndPoint.Y));
                        break;
                    case "LWPOLYLINE":
                        Polyline pLine = (Polyline)oid.GetObject(OpenMode.ForRead);
                        ParsePolyLine(pLine, baseX, baseY, xScale, yScale);
                        break;
                    case "INSERT":
                        DBObject obj = oid.GetObject(OpenMode.ForRead);
                        if (obj is BlockReference)
                        {
                            BlockReference br = obj as BlockReference;
                            BlockTableRecord btr = br.BlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                            ParseObjects(btr, baseX + br.Position.X, baseY + br.Position.Y, xScale * br.ScaleFactors.X, yScale * br.ScaleFactors.Y);
                        }
                        break;
                    default:
                        //System.Windows.Forms.MessageBox.Show(oid.ObjectClass.DxfName);
                        break;
                }
            }
        }

        /// <summary>
        /// 判断Text是否已存在（重复的Text）
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private bool TextExists(AcText text)
        {
            foreach (AcText t in _texts)
            {
                if (t.Value == text.Value && t.Position == text.Position && t.Height == text.Height)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 排序识别到的对象，按从上到下从左到右的顺序
        /// </summary>
        private void SortObjects()
        {
            if (_tableHLines.Count == 0 || _tableVLines.Count == 0)
            {
                throw new AcTableParseException("未能识别到表格");
            }
            if (_texts.Count == 0)
            {
                throw new AcTableParseException("未能识别到表格内容");
            }
            _texts.Sort((item1, item2) => Math.Sign(item1.Position.X - item2.Position.X));
            //上面的线Y坐标更大，所以反着排序
            _tableHLines.Sort((item1, item2) => Math.Sign(item2.XorY - item1.XorY));
            _tableVLines.Sort((item1, item2) => Math.Sign(item1.XorY - item2.XorY));
        }

        /// <summary>
        /// 解析多段线
        /// </summary>
        /// <param name="pLine"></param>
        private void ParsePolyLine(Polyline pLine, double baseX = 0, double baseY = 0, double xScale = 1, double yScale = 1)
        {
            Point2d pt0 = pLine.GetPoint2dAt(0);
            Point2d ptPre = pt0;
            for (int i = 1; i < pLine.NumberOfVertices; i++)
            {
                Point2d ptCur = pLine.GetPoint2dAt(i);
                ParseLine(new Point2d(baseX + xScale * ptPre.X, baseY + yScale * ptPre.Y),
                    new Point2d(baseX + xScale * ptCur.X, baseY + yScale * ptCur.Y));

                if (pLine.Closed && i == pLine.NumberOfVertices - 1)
                {
                    ParseLine(new Point2d(baseX + xScale * ptCur.X, baseY + yScale * ptCur.Y),
                        new Point2d(baseX + xScale * pt0.X, baseY + yScale * pt0.Y));
                }
                else
                {
                    ptPre = ptCur;
                }
            }
        }

        /// <summary>
        /// 解析直线
        /// </summary>
        /// <param name="line"></param>
        private void ParseLine(Point2d ptStart, Point2d ptEnd)
        {
            if (!CommandUtils.IsOrthogonalLine(ptStart, ptEnd))
            {
                //表格里会有斜线代表空格，直接跳过
                //throw new AcTableParseException(string.Format("只支持正交的直线 {0:s} {1:s}", ptStart.ToString(), ptEnd.ToString()));
                return;
            }
            List<AcTableLine> list;
            if (CommandUtils.Compare(ptStart.X, ptEnd.X) == 0)
            {
                list = _tableVLines;
            }
            else
            {
                list = _tableHLines;
            }
            bool add = false;
            foreach (AcTableLine tl in list)
            {
                if (tl.AcceptSegment(ptStart, ptEnd))
                {
                    tl.AddSegment(ptStart, ptEnd);
                    add = true;
                    break;
                }
            }
            if (!add)
            {
                list.Add(new AcTableLine(ptStart, ptEnd));
            }
        }


    }

}
