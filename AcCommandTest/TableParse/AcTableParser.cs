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
        /// <summary>
        /// 文本对象的结构
        /// </summary>
        struct Text
        {
            private Point2d position;
            /// <summary>
            /// 文本位置（中心点）
            /// </summary>
            public Point2d Position
            {
                get { return position; }
            }

            private double height;
            /// <summary>
            /// 文本高度
            /// </summary>
            public double Height
            {
                get { return height; }
            }


            private string value;
            /// <summary>
            /// 文本值
            /// </summary>
            public string Value
            {
                get { return value; }
            }

            public Text(string t, Point2d pos, double height)
            {
                this.value = t;
                this.position = pos;
                this.height = height;
            }
        }

        /// <summary>
        /// 格线方向
        /// </summary>
        enum TableLineDirection
        {
            H,
            V
        }

        /// <summary>
        /// 线段
        /// </summary>
        struct LineSegment
        {
            private Point2d _start;
            /// <summary>
            /// 开始点
            /// </summary>
            public Point2d Start
            {
                get { return _start; }
            }

            private Point2d _end;
            /// <summary>
            /// 结束点
            /// </summary>
            public Point2d End
            {
                get { return _end; }
            }

            public LineSegment(Point2d start, Point2d end)
            {
                _start = start;
                _end = end;
            }
        }

        /// <summary>
        /// 表格线的结构
        /// </summary>
        class TableLine
        {
            private TableLineDirection _direction;

            private List<LineSegment> _segments = new List<LineSegment>();
            /// <summary>
            /// 格线中的线段
            /// </summary>
            public List<LineSegment> Segments { get { return _segments; } }

            private double _xory;
            /// <summary>
            /// 线的位置，对于横线来说是Y值，对于竖线来说是X值
            /// </summary>
            public double XorY { get { return _xory; } }

            /// <summary>
            /// 是否接受一个线段，即该线段是否可以归到这条格线中
            /// 这里默认线段已经是正交的，应在调用之前做好判断
            /// </summary>
            /// <param name="start"></param>
            /// <param name="end"></param>
            /// <returns></returns>
            public bool AcceptSegment(Point2d pt1, Point2d pt2)
            {
                TableLineDirection direction;
                double xory = 0;
                if (DoubleValueCompare(pt1.X, pt2.X) == 0)
                {
                    direction = TableLineDirection.V;
                    xory = pt1.X;
                }
                else
                {
                    direction = TableLineDirection.H;
                    xory = pt1.Y;
                }
                return direction == _direction && (_segments.Count == 0 || DoubleValueCompare(xory, _xory) == 0);
            }

            /// <summary>
            /// 向格线中添加一段
            /// 应提前调用Accept方法判断是否应该属于该格线，本方法中不再判断
            /// </summary>
            /// <param name="pt1"></param>
            /// <param name="pt2"></param>
            public void AddSegment(Point2d pt1, Point2d pt2)
            {
                Point2d start, end;
                if (pt1.Y > pt2.Y || pt1.X < pt2.X)
                {
                    start = pt1;
                    end = pt2;
                }
                else
                {
                    start = pt2;
                    end = pt1;
                }

                if (_direction == TableLineDirection.H)
                {
                    if (_segments.Count == 0)
                    {
                        _xory = start.Y;
                        _segments.Add(new LineSegment(start, end));
                    }
                    for (int i = 0; i < _segments.Count; i++)
                    {
                        if (start.X < _segments[i].Start.X)
                        {
                            _segments.Insert(i, new LineSegment(start, end));
                            break;
                        }
                    }
                }
                else
                {
                    if (_segments.Count == 0)
                    {
                        _xory = start.X;
                        _segments.Add(new LineSegment(start, end));
                    }
                    for (int i = 0; i < _segments.Count; i++)
                    {
                        if (start.Y > _segments[i].Start.Y)
                        {
                            _segments.Insert(i, new LineSegment(start, end));
                            break;
                        }
                    }
                }
            }

            public TableLine(TableLineDirection direction)
            {
                _direction = direction;
            }

            /// <summary>
            /// 构造器
            /// 这里默认线段已经是正交的，应在调用之前做好判断
            /// </summary>
            /// <param name="start"></param>
            /// <param name="end"></param>
            public TableLine(Point2d pt1, Point2d pt2)
            {
                if (DoubleValueCompare(pt1.X, pt2.X) == 0)
                {
                    _direction = TableLineDirection.V;
                    _xory = pt1.X;
                }
                else
                {
                    _direction = TableLineDirection.H;
                    _xory = pt2.Y;
                }
                Point2d start, end;
                if (pt1.Y > pt2.Y || pt1.X < pt2.X)
                {
                    start = pt1;
                    end = pt2;
                }
                else
                {
                    start = pt2;
                    end = pt1;
                }
                _segments.Add(new LineSegment(start, end));
            }
        }

        /// <summary>
        /// 格子
        /// </summary>
        class TableCell
        {
            public int Row { get; set; }
            public int Col { get; set; }
            public int RowSpan { get; set; }
            public int ColSpan { get; set; }
            public Point2d TopLeft { get; set; }
            public Point2d BottomRight { get; set; }
            public List<Text> Texts { get; set; }
            private string _cellValue;
            public string CellValue { get { return _cellValue; } }
            public string MergedCellValue { get { return _cellValue; } }

            public TableCell()
            {
                RowSpan = 1;
                ColSpan = 1;
                Texts = new List<Text>();
                _cellValue = "";
            }

            /// <summary>
            /// 计算文本值，处理多个Text的合并，处理特殊情况
            /// </summary>
            public void CalcCellValue()
            {
                if (Texts.Count == 0)
                {
                    return;
                }
                else if (Texts.Count == 1)
                {
                    _cellValue = ProcessSpecialText(Texts[0].Value);
                    return;
                }
                Texts.Sort((text1, text2) => (int)(text1.Position.X - text2.Position.X + text2.Position.Y - text1.Position.Y));
                List<string> parts = new List<string>();
                foreach (Text text in Texts)
                {
                    parts.Add(ProcessSpecialText(text.Value));
                }
                //特殊处理㎡和m³，在有的表格中会分开两个文本对象
                for (int i = 0; i < parts.Count; i++)
                {
                    if (parts[i] == "2" || parts[i] == "3")
                    {
                        for (int j = 0; j < parts.Count; j++)
                        {
                            //带不带m
                            if (!parts[j].Contains('m'))
                            {
                                continue;
                            }
                            //是否在2或3的下方
                            if (Texts[j].Position.Y >= Texts[i].Position.Y)
                            {
                                continue;
                            }
                            //纵向位置不能差的太远
                            if (Texts[i].Position.Y - Texts[j].Position.Y > (Texts[i].Height + Texts[j].Height) / 2)
                            {
                                continue;
                            }
                            if (parts[i] == "2")
                            {
                                parts[j] = parts[j].Replace("m", "㎡");
                            }
                            else if (parts[i] == "3")
                            {
                                parts[j] = parts[j].Replace("m", "m³");
                            }
                            parts[i] = "";
                            break;
                        }
                    }
                }
                int prevIndex = 0;
                for (int i = 0; i < parts.Count; i++)
                {
                    if (parts[i] != "")
                    {
                        //判断是否需要加分隔
                        if (i > 0 && _cellValue != "" && Texts[prevIndex].Position.Y - Texts[i].Position.Y > (Texts[i].Height + Texts[prevIndex].Height) / 2)
                        {
                            _cellValue += " ";
                        }
                        _cellValue += parts[i];
                        prevIndex = i;
                    }
                }
            }

            /// <summary>
            /// 处理特殊文本
            /// </summary>
            /// <param name="s"></param>
            /// <returns></returns>
            private string ProcessSpecialText(string s)
            {
                return s.Replace("m3/", "m³"); //有的多行文本中的立方米，解析出来是奇怪的m3/三个字符
            }
        }

        private List<Text> _texts = new List<Text>();
        private List<TableLine> _tableHLines = new List<TableLine>();
        private List<TableLine> _tableVLines = new List<TableLine>();
        private TableCell[,] _cells;
        private int _rowCount;
        private int _colCount;

        /// <summary>
        /// 解析选中对象为一个表格，如果无法识别会抛出AcTableParseException异常，需调用方处理异常
        /// </summary>
        /// <param name="objects"></param>
        /// <returns></returns>
        public static AcTable ParseTable(IEnumerable objects)
        {
            return new AcTableParser().Parse(objects);
        }

        /// <summary>
        /// 解析表格
        /// </summary>
        /// <param name="objects"></param>
        /// <returns></returns>
        private AcTable Parse(IEnumerable objects)
        {
            ParseObjects(objects);
            BuildCells();
            PutTextToCells();
            CalcCellsValue();
            return GenerateResult();
        }

        /// <summary>
        /// 生成AcTable类型的返回结果
        /// </summary>
        /// <returns></returns>
        private AcTable GenerateResult()
        {
            AcTable table = new AcTable();
            table.Cells = new AcTableCell[_rowCount][];
            for (int i = 0; i < _rowCount; i++)
            {
                table.Cells[i] = new AcTableCell[_colCount];
                for (int j = 0; j < _colCount; j++)
                {
                    table.Cells[i][j] = new AcTableCell();
                    table.Cells[i][j].Row = i;
                    table.Cells[i][j].Col = j;
                    table.Cells[i][j].Value = _cells[i, j].CellValue;
                }
            }

            return table;
        }

        /// <summary>
        /// 计算每个格子的文本值
        /// </summary>
        private void CalcCellsValue()
        {
            foreach (TableCell cell in _cells)
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
            foreach (Text text in _texts)
            {
                bool put = false;
                for (int j = startCol; j < _colCount; j++)
                {
                    for (int i = 0; i < _rowCount; i++)
                    {
                        if (PointInRect(text.Position, _cells[i, j].TopLeft, _cells[i, j].BottomRight))
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
            _cells = new TableCell[_rowCount, _colCount];
            for (int i = 0; i < _tableHLines.Count - 1; i++)
            {
                for (int j = 0; j < _tableVLines.Count - 1; j++)
                {
                    TableCell tc = new TableCell();
                    tc.Row = i;
                    tc.Col = j;
                    tc.TopLeft = new Point2d(_tableVLines[j].XorY, _tableHLines[i].XorY);
                    tc.BottomRight = new Point2d(_tableVLines[j + 1].XorY, _tableHLines[i + 1].XorY);
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
                        _texts.Add(new Text(text.TextString, new Point2d(text.Position.X, text.Position.Y), text.Height));
                        break;
                    case "MTEXT":
                        MText mText = (MText)oid.GetObject(OpenMode.ForRead);
                        _texts.Add(new Text(mText.Text, new Point2d(mText.Location.X, mText.Location.Y), mText.Height));
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
            if (!IsOrthogonalLine(ptStart, ptEnd))
            {
                //表格里会有斜线代表空格，直接跳过
                //throw new AcTableParseException(string.Format("只支持正交的直线 {0:s} {1:s}", ptStart.ToString(), ptEnd.ToString()));
                return;
            }
            List<TableLine> list;
            if (DoubleValueCompare(ptStart.X, ptEnd.X) == 0)
            {
                list = _tableVLines;
            }
            else
            {
                list = _tableHLines;
            }
            bool add = false;
            foreach (TableLine tl in list)
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
                list.Add(new TableLine(ptStart, ptEnd));
            }
        }

        /// <summary>
        /// 直线是否正交模式（横或竖）
        /// </summary>
        /// <param name="pt1"></param>
        /// <param name="pt2"></param>
        /// <returns></returns>
        private bool IsOrthogonalLine(Point2d pt1, Point2d pt2)
        {
            return DoubleValueCompare(pt1.X, pt2.X) == 0 || DoubleValueCompare(pt1.Y, pt2.Y) == 0;
        }

        /// <summary>
        /// 浮点数比较（Round后比较）
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns>1,0,-1</returns>
        private static int DoubleValueCompare(double d1, double d2)
        {
            return Math.Sign(Math.Round(d1 - d2));
        }

        /// <summary>
        /// 判断点是否在区域里
        /// </summary>
        /// <param name="point"></param>
        /// <param name="rectTopLeft"></param>
        /// <param name="rectBottomRight"></param>
        /// <returns></returns>
        private bool PointInRect(Point2d point, Point2d rectTopLeft, Point2d rectBottomRight)
        {
            return point.X > rectTopLeft.X && point.X < rectBottomRight.X
                && point.Y < rectTopLeft.Y && point.Y > rectBottomRight.Y;
        }

        /// <summary>
        /// 判断某一行是否全空
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private bool IsRowEmpty(int row)
        {
            for (int j = 0; j < _colCount; j++)
            {
                if (_cells[row, j].CellValue != "")
                {
                    return false;
                }
            }
            return true;
        }
    }

}
