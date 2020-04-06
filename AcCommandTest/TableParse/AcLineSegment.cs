using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 线段
    /// </summary>
    struct AcLineSegment
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

        public AcLineSegment(Point2d start, Point2d end)
        {
            _start = start;
            _end = end;
        }
    }

}
