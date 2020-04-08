using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 表格线的结构
    /// </summary>
    class AcTableLine
    {
        private AcTableLineDirection _direction;

        private List<AcLineSegment> _segments = new List<AcLineSegment>();
        /// <summary>
        /// 格线中的线段
        /// </summary>
        public List<AcLineSegment> Segments { get { return _segments; } }

        private double _xory;
        /// <summary>
        /// 线的位置，对于横线来说是Y值，对于竖线来说是X值
        /// </summary>
        public double XorY { get { return _xory; } }

        public Point2d _start;
        /// <summary>
        /// 开始位置
        /// </summary>
        public Point2d Start { get { return _start; } }

        public Point2d _end;
        /// <summary>
        /// 结束位置
        /// </summary>
        public Point2d End { get { return _end; } }

        /// <summary>
        /// 长度
        /// </summary>
        public double Length
        {
            get
            {
                if (_direction == AcTableLineDirection.H)
                {
                    return End.X - Start.X;
                }
                else
                {
                    return Start.Y - End.Y;
                }
            }
        }

        /// <summary>
        /// 是否接受一个线段，即该线段是否可以归到这条格线中
        /// 这里默认线段已经是正交的，应在调用之前做好判断
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public bool AcceptSegment(Point2d pt1, Point2d pt2)
        {
            AcTableLineDirection direction;
            double xory = 0;
            if (CommandUtils.Compare(pt1.X, pt2.X) == 0)
            {
                direction = AcTableLineDirection.V;
                xory = pt1.X;
            }
            else
            {
                direction = AcTableLineDirection.H;
                xory = pt1.Y;
            }
            return direction == _direction && (_segments.Count == 0 || CommandUtils.Compare(xory, _xory) == 0);
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
            if (CommandUtils.Compare(pt1.Y, pt2.Y) > 0 || CommandUtils.Compare(pt1.X, pt2.X) < 0)
            {
                start = pt1;
                end = pt2;
            }
            else
            {
                start = pt2;
                end = pt1;
            }

            if (_direction == AcTableLineDirection.H)
            {
                if (_segments.Count == 0)
                {
                    _xory = start.Y;
                    _segments.Add(new AcLineSegment(start, end));
                }
                bool add = false;
                for (int i = 0; i < _segments.Count; i++)
                {
                    if (start.X < _segments[i].Start.X)
                    {
                        _segments.Insert(i, new AcLineSegment(start, end));
                        add = true;
                        break;
                    }
                }
                if (!add)
                {
                    _segments.Add(new AcLineSegment(start, end));
                }
                if (_start.X > start.X)
                {
                    _start = start;
                }
                if (_end.X < end.X)
                {
                    _end = end;
                }
            }
            else
            {
                if (_segments.Count == 0)
                {
                    _xory = start.X;
                    _segments.Add(new AcLineSegment(start, end));
                }
                bool add = false;
                for (int i = 0; i < _segments.Count; i++)
                {
                    if (start.Y > _segments[i].Start.Y)
                    {
                        _segments.Insert(i, new AcLineSegment(start, end));
                        add = true;
                        break;
                    }
                }
                if (!add)
                {
                    _segments.Add(new AcLineSegment(start, end));
                }
                if (_start.Y < start.Y)
                {
                    _start = start;
                }
                if (_end.Y > end.Y)
                {
                    _end = end;
                }
            }
        }

        /// <summary>
        /// 在某个点上有无线段
        /// </summary>
        /// <param name="xory"></param>
        /// <returns></returns>
        public bool HasSegmentOn(double xory)
        {
            foreach (AcLineSegment seg in _segments)
            {
                if (_direction == AcTableLineDirection.H)
                {
                    if (seg.Start.X <= xory && seg.End.X >= xory)
                    {
                        return true;
                    }
                }
                else
                {
                    if (seg.Start.Y >= xory && seg.End.Y <= xory)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 构造器
        /// 这里默认线段已经是正交的，应在调用之前做好判断
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public AcTableLine(Point2d pt1, Point2d pt2)
        {
            if (CommandUtils.Compare(pt1.X, pt2.X) == 0)
            {
                _direction = AcTableLineDirection.V;
                _xory = pt1.X;
            }
            else
            {
                _direction = AcTableLineDirection.H;
                _xory = pt2.Y;
            }
            Point2d start, end;
            if (CommandUtils.Compare(pt1.Y, pt2.Y) > 0 || CommandUtils.Compare(pt1.X, pt2.X) < 0)
            {
                start = pt1;
                end = pt2;
            }
            else
            {
                start = pt2;
                end = pt1;
            }
            _start = start;
            _end = end;
            _segments.Add(new AcLineSegment(start, end));
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("direction: " + _direction.ToString());
            sb.AppendLine("xory: " + XorY.ToString());
            sb.AppendLine("Start: " + string.Format("( {0:f1}, {1:f1} )", Start.X, Start.Y));
            sb.AppendLine("End: " + string.Format("( {0:f1}, {1:f1} )", End.X, End.Y));
            sb.AppendLine("Segments: ");
            foreach (var seg in _segments)
            {
                sb.AppendLine(string.Format("    ( {0:f1}, {1:f1} ) - ( {2:f1}, {3:f1} )", seg.Start.X, seg.Start.Y, seg.End.X, seg.End.Y));
            }
            return string.Format(sb.ToString());
        }
    }
}
