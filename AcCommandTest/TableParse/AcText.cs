using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 文本对象的结构
    /// </summary>
    struct AcText
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

        public AcText(string t, Point2d pos, double height)
        {
            this.value = t;
            this.position = pos;
            this.height = height;
        }
    }
}
