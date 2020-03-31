using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AcCommandTest
{
    /// <summary>
    /// 识别表格的异常类
    /// </summary>
    public class AcTableParseException : Exception
    {
        public AcTableParseException(string message) : base(message) { }
        public AcTableParseException(string message, Exception innerException) : base(message, innerException) { }
    }
}
