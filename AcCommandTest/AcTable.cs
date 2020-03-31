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
        /// 是否有列头
        /// </summary>
        public bool HasHeader { get; set; }
        /// <summary>
        /// 是否有合计行
        /// </summary>
        public bool HasFooter { get; set; }
        /// <summary>
        /// 列的数量
        /// </summary>
        public int ColCount { get; set; }
        /// <summary>
        /// 数据行的数量（不算列头和合计行）
        /// </summary>
        public int DataRowCount { get; set; }
        /// <summary>
        /// 列头
        /// </summary>
        public string[] Header { get; set; }
        /// <summary>
        /// 合计行
        /// </summary>
        public string[] Footer { get; set; }
        /// <summary>
        /// 数据
        /// </summary>
        public string[][] Data { get; set; }

        public override string ToString()
        {
            return string.Format("ColCount: {0:d}, DataRowCount: {1:d}, HasHeader: {2:s}, HasFooter: {3:s}",
                ColCount, DataRowCount, HasHeader.ToString(), HasFooter.ToString());
        }
    }
}
