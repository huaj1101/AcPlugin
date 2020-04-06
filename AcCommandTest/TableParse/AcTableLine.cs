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
            if (CommandUtils.DoubleValueCompare(pt1.X, pt2.X) == 0)
            {
                direction = AcTableLineDirection.V;
                xory = pt1.X;
            }
            else
            {
                direction = AcTableLineDirection.H;
                xory = pt1.Y;
            }
            return direction == _direction && (_segments.Count == 0 || CommandUtils.DoubleValueCompare(xory, _xory) == 0);
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

            if (_direction == AcTableLineDirection.H)
            {
                if (_segments.Count == 0)
                {
                    _xory = start.Y;
                    _segments.Add(new AcLineSegment(start, end));
                }
                for (int i = 0; i < _segments.Count; i++)
                {
                    if (start.X < _segments[i].Start.X)
                    {
                        _segments.Insert(i, new AcLineSegment(start, end));
                        break;
                    }
                }
            }
            else
            {
                if (_segments.Count == 0)
                {
                    _xory = start.X;
                    _segments.Add(new AcLineSegment(start, end));
                }
                for (int i = 0; i < _segments.Count; i++)
                {
                    if (start.Y > _segments[i].Start.Y)
                    {
                        _segments.Insert(i, new AcLineSegment(start, end));
                        break;
                    }
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

        public AcTableLine(AcTableLineDirection direction)
        {
            _direction = direction;
        }

        /// <summary>
        /// 构造器
        /// 这里默认线段已经是正交的，应在调用之前做好判断
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public AcTableLine(Point2d pt1, Point2d pt2)
        {
            if (CommandUtils.DoubleValueCompare(pt1.X, pt2.X) == 0)
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
            _segments.Add(new AcLineSegment(start, end));
        }
    }
}
