using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
        private static Dictionary<string, string> _specialChars = new Dictionary<string, string>();
        private static Dictionary<string, string> _specialStrs = new Dictionary<string, string>();
        private static Regex _commonSpecialCharRegex;

        static AcTableCell()
        {
            //初始化特殊字符对照表
            _specialChars["%%P"] = "±";
            _specialChars["%%p"] = "±";
            _specialChars["%%C"] = "φ";
            _specialChars["%%c"] = "φ";
            _specialChars["%%D"] = "°";
            _specialChars["%%d"] = "°";
            _specialChars["%%%"] = "%";
            _specialChars["%%130"] = "φ";
            _specialChars["%%131"] = "φ";
            _specialChars["%%132"] = "φ";
            _specialChars["%%133"] = "φ";
            _specialChars["m3/"] = "m³";
            _commonSpecialCharRegex = new Regex(@"%%\d{3}");

            _specialStrs["m ) 3"] = "m³)";
            _specialStrs["m) 3"] = "m³)";
            _specialStrs["m 3"] = "m³";
            _specialStrs["m³ )"] = "m³)";
        }

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
            Texts.Sort((text1, text2) => (int)((text1.Position.X - text2.Position.X + 2 * (text2.Position.Y - text1.Position.Y)) * 100));
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
                    //判断是否需要加换行
                    if (i > 0 && InnerCell.Value != "" && Texts[prevIndex].Position.Y - Texts[i].Position.Y > (Texts[i].Height + Texts[prevIndex].Height) / 2)
                    {
                        InnerCell.Value += "\n";
                    }
                    else if (InnerCell.Value != "")
                    {
                        InnerCell.Value += " ";
                    }
                    InnerCell.Value += parts[i];
                    prevIndex = i;
                }
            }
            foreach (string key in _specialStrs.Keys)
            {
                InnerCell.Value = InnerCell.Value.Replace(key, _specialStrs[key]);
            }
        }

        /// <summary>
        /// 处理特殊文本
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private string ProcessSpecialText(string s)
        {
            foreach (string key in _specialChars.Keys)
            {
                s = s.Replace(key, _specialChars[key]);
            }
            s = _commonSpecialCharRegex.Replace(s, "");
            return s.Trim();
        }
    }
}
