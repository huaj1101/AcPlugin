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
            BuildCells();
            PutTextToCells();
            CalcCellMerge();
            CalcCellsValue();
            return GenerateResult();
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
                            !_cells[i, j].RightLine.HasSegmentOn(_cells[i, j].Center.Y))
                        {
                            _cells[i, j].InnerCell.MasterCell.ColSpan += 1;
                            for (int row = i; row < i + _cells[i, j].InnerCell.MasterCell.RowSpan; row++)
                            {
                                _cells[row, j + 1].InnerCell.MasterCell = _cells[i, j].InnerCell.MasterCell;
                                _cells[row, j + 1].InnerCell.CellType = TableCellType.MergedSlave;
                            }
                        }
                        else if (_cells[i, j].InnerCell.Col == _cells[i, j].InnerCell.MasterCell.Col &&
                            !_cells[i, j].BottomLine.HasSegmentOn(_cells[i, j].Center.X))
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
                        if (!_cells[i, j].RightLine.HasSegmentOn(_cells[i, j].Center.Y))
                        {
                            _cells[i, j].InnerCell.CellType = TableCellType.MergedMaster;
                            _cells[i, j].InnerCell.ColSpan = 2;
                            _cells[i, j + 1].InnerCell.MasterCell = _cells[i, j].InnerCell;
                            _cells[i, j + 1].InnerCell.CellType = TableCellType.MergedSlave;
                        }
                        if (!_cells[i, j].BottomLine.HasSegmentOn(_cells[i, j].Center.X))
                        {
                            _cells[i, j].InnerCell.CellType = TableCellType.MergedMaster;
                            _cells[i, j].InnerCell.RowSpan = 2;
                            _cells[i + 1, j].InnerCell.MasterCell = _cells[i, j].InnerCell;
                            _cells[i + 1, j].InnerCell.CellType = TableCellType.MergedSlave;
                            if (_cells[i, j].InnerCell.ColSpan == 2)
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
        private void ParseObjects(IEnumerable objects)
        {
            foreach (ObjectId oid in objects)
            {
                switch (oid.ObjectClass.DxfName)
                {
                    case "TEXT":
                        DBText text = (DBText)oid.GetObject(OpenMode.ForRead);
                        _texts.Add(new AcText(text.TextString, new Point2d(text.Position.X, text.Position.Y), text.Height));
                        break;
                    case "MTEXT":
                        MText mText = (MText)oid.GetObject(OpenMode.ForRead);
                        _texts.Add(new AcText(mText.Text, new Point2d(mText.Location.X, mText.Location.Y), mText.Height));
                        break;
                    case "LINE":
                        Line line = (Line)oid.GetObject(OpenMode.ForRead);
                        ParseLine(new Point2d(line.StartPoint.X, line.StartPoint.Y), new Point2d(line.EndPoint.X, line.EndPoint.Y));
                        break;
                    case "LWPOLYLINE":
                        Polyline pLine = (Polyline)oid.GetObject(OpenMode.ForRead);
                        ParsePolyLine(pLine);
                        break;
                    default:
                        //System.Windows.Forms.MessageBox.Show(oid.ObjectClass.DxfName);
                        break;
                }
            }
            if (_tableHLines.Count == 0 || _tableVLines.Count == 0)
            {
                throw new AcTableParseException("未能识别到表格");
            }
            if (_texts.Count == 0)
            {
                throw new AcTableParseException("未能识别到表格内容");
            }
            _texts.Sort((item1, item2) => (int)(item1.Position.X - item2.Position.X));
            //上面的线Y坐标更大，所以反着排序
            _tableHLines.Sort((item1, item2) => (int)(item2.XorY - item1.XorY));
            _tableVLines.Sort((item1, item2) => (int)(item1.XorY - item2.XorY));
        }

        /// <summary>
        /// 解析多段线
        /// </summary>
        /// <param name="pLine"></param>
        private void ParsePolyLine(Polyline pLine)
        {
            Point2d pt0 = pLine.GetPoint2dAt(0);
            Point2d ptPre = pt0;
            for (int i = 1; i < pLine.NumberOfVertices; i++)
            {
                Point2d ptCur = pLine.GetPoint2dAt(i);
                ParseLine(ptPre, ptCur);

                if (pLine.Closed && i == pLine.NumberOfVertices - 1)
                {
                    ParseLine(ptCur, pt0);
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
            if (CommandUtils.DoubleValueCompare(ptStart.X, ptEnd.X) == 0)
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
