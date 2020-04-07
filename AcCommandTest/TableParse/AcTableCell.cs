using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 格子
    /// </summary>
    class AcTableCell
    {
        public Point2d TopLeft { get; set; }
        public Point2d BottomRight { get; set; }
        public Point2d Center { get; set; }
        public AcTableLine LeftLine { get; set; }
        public AcTableLine RightLine { get; set; }
        public AcTableLine TopLine { get; set; }
        public AcTableLine BottomLine { get; set; }
        public List<AcText> Texts { get; set; }
        public TableCell InnerCell { get; set; }

        public AcTableCell()
        {
            Texts = new List<AcText>();
            InnerCell = new TableCell();
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
                InnerCell.Value = ProcessSpecialText(Texts[0].Value);
                return;
            }
            Texts.Sort((text1, text2) => (int)(text1.Position.X - text2.Position.X + text2.Position.Y - text1.Position.Y));
            List<string> parts = new List<string>();
            foreach (AcText text in Texts)
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
                    if (i > 0 && InnerCell.Value != "" && Texts[prevIndex].Position.Y - Texts[i].Position.Y > (Texts[i].Height + Texts[prevIndex].Height) / 2)
                    {
                        InnerCell.Value += "\n";
                    }
                    InnerCell.Value += parts[i];
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
}
