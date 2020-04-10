using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    static class CommandUtils
    {
        /// <summary>
        /// 直线是否正交模式（横或竖）
        /// </summary>
        /// <param name="pt1"></param>
        /// <param name="pt2"></param>
        /// <returns></returns>
        public static bool IsOrthogonalLine(Point2d pt1, Point2d pt2)
        {
            return Compare(pt1.X, pt2.X) == 0 || Compare(pt1.Y, pt2.Y) == 0;
        }

        /// <summary>
        /// 比较浮点数（Round后比较）
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns>-1, 0, 1</returns>
        public static int Compare(double d1, double d2, int decimals = 3)
        {
            return Math.Sign(Math.Round(d1 - d2, decimals));
        }

        /// <summary>
        /// 判断点是否在区域里
        /// </summary>
        /// <param name="point"></param>
        /// <param name="rectTopLeft"></param>
        /// <param name="rectBottomRight"></param>
        /// <returns></returns>
        public static bool PointInRect(Point2d point, Point2d rectTopLeft, Point2d rectBottomRight)
        {
            return point.X >= rectTopLeft.X && point.X <= rectBottomRight.X
                && point.Y <= rectTopLeft.Y && point.Y >= rectBottomRight.Y;
        }

        /// <summary>
        /// 获取一个实体的平面中心点
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public static Point2d GetCenterPoint(Entity entity)
        {
            Extents3d e3d = entity.GeometricExtents;
            return new Point2d(
                (e3d.MinPoint.X + e3d.MaxPoint.X) / 2,
                (e3d.MinPoint.Y + e3d.MaxPoint.Y) / 2
                );
        }
    }
}
