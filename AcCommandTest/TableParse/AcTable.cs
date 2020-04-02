using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 解析AutoCad表格的结果
    /// </summary>
    public class AcTable
    {
        /// <summary>
        /// 行的数量
        /// </summary>
        public int RowCount { get { return Cells == null ? 0 : Cells.Length; } }
        /// <summary>
        /// 列的数量
        /// </summary>
        public int ColCount { get { return Cells == null ? 0 : Cells[0].Length; } }
        /// <summary>
        /// 数据
        /// </summary>
        public AcTableCell[][] Cells { get; set; }

        public override string ToString()
        {
            return string.Format("RowCount: {0:d}, ColCount: {1:d}", RowCount, ColCount);
        }
    }

    /// <summary>
    /// AutoCad表格的格子
    /// </summary>
    public class AcTableCell
    {
        /// <summary>
        /// 行
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// 列
        /// </summary>
        public int Col { get; set; }
        /// <summary>
        /// 跨行数
        /// </summary>
        public int RowSpan { get; set; }
        /// <summary>
        /// 跨列数
        /// </summary>
        public int ColSpan { get; set; }
        /// <summary>
        /// 格子类型
        /// </summary>
        public AcTableCellType CellType { get; set; }
        /// <summary>
        /// 主格，CellType为MergedSlave时适用
        /// </summary>
        public AcTableCell MasterCell { get; set; }
        /// <summary>
        /// 格子的值
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// AutoCad表格的格子类型
    /// </summary>
    public enum AcTableCellType
    {
        /// <summary>
        /// 正常格子（无合并）
        /// </summary>
        Normal,
        /// <summary>
        /// 合并格子中的主格（左上角）
        /// </summary>
        MergedMaster,
        /// <summary>
        /// 合并格子中的从格
        /// </summary>
        MergedSlave
    }
}
