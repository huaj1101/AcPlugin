using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcCommandTest
{
    public class AcTableParser
    {
        struct CellValue
        {
            private Point2d position;
            public Point2d Position
            {
                get { return position; }
            }

            private string text;
            public string Text
            {
                get { return text; }
            }

            public CellValue(Point2d pos, string t)
            {
                this.position = pos;
                this.text = t;
            }
        }

        struct TableLine
        {
            private Point2d point1;
            public Point2d Point1
            {
                get { return point1; }
            }

            private Point2d point2;
            public Point2d Point2
            {
                get { return point2; }
            }
            public TableLine(Point2d pt1, Point2d pt2)
            {
                this.point1 = pt1;
                this.point2 = pt2;
            }
        }

        private List<CellValue> _cellValues = new List<CellValue>();
        private List<TableLine> _tableLines = new List<TableLine>();

        public static AcTable ParseTable(SelectionSet objects)
        {
            return new AcTableParser().Parse(objects);
        }

        private AcTable Parse(SelectionSet objects)
        {
            foreach (SelectedObject so in objects)
            {
                switch (so.ObjectId.ObjectClass.DxfName)
                {
                    case "TEXT":
                        DBText text = (DBText)so.ObjectId.GetObject(OpenMode.ForRead);
                        _cellValues.Add(new CellValue(new Point2d(text.Position.X, text.Position.Y), text.TextString));
                        break;
                    case "MTEXT":
                        MText mtext = (MText)so.ObjectId.GetObject(OpenMode.ForRead);
                        _cellValues.Add(new CellValue(new Point2d(mtext.Location.X, mtext.Location.Y), mtext.Contents));
                        break;
                    case "LINE":
                        Line line = (Line)so.ObjectId.GetObject(OpenMode.ForRead);
                        _tableLines.Add(new TableLine(new Point2d(line.StartPoint.X, line.StartPoint.Y), new Point2d(line.EndPoint.X, line.EndPoint.Y)));
                        break;
                    case "LWPOLYLINE":
                        Polyline pLine = (Polyline)so.ObjectId.GetObject(OpenMode.ForRead);
                        Point2d pt0 = pLine.GetPoint2dAt(0);
                        Point2d ptPre = pt0;
                        for (int i = 1; i < pLine.NumberOfVertices; i++)
                        {
                            Point2d ptCur = pLine.GetPoint2dAt(i);
                            if (ptCur.X != ptPre.X && ptCur.Y != ptPre.Y)
                            {
                                throw new Exception("只支持由正交的直线组成的多段线");
                            }
                            _tableLines.Add(new TableLine(ptPre, ptCur));

                            if (pLine.Closed && i == pLine.NumberOfVertices - 1)
                            {
                                if (ptCur.X != pt0.X && ptCur.Y != pt0.Y)
                                {
                                    throw new Exception("只支持由正交的直线组成的多段线");
                                }
                                _tableLines.Add(new TableLine(ptCur, pt0));
                            }
                            else
                            {
                                ptPre = ptCur;
                            }
                        }
                        break;
                    default:
                        //System.Windows.Forms.MessageBox.Show(so.ObjectId.ObjectClass.DxfName);
                        break;
                }
            }
            AcTable table = new AcTable();
            table.ColCount = _tableLines.Count;
            table.DataRowCount = _cellValues.Count;
            return table;
        }
    }


    public class AcTable
    {
        public bool HasHeader { get; set; }
        public bool HasFooter { get; set; }
        public int ColCount { get; set; }
        public int DataRowCount { get; set; }
        public string[] Header { get; set; }
        public string[] Footer { get; set; }
        public string[,] Data { get; set; }

        public static AcTable Empty { get { return new AcTable(); } }

        public override string ToString()
        {
            return string.Format("LineCount {0:d}, CellCount {1:d}", ColCount, DataRowCount);
        }
    }
}
