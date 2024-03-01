using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0
{
    public class NRAFormatException : Exception
    {
        public NRAFormatException() { }

        public NRAFormatException(string message) : base(message) { }
    }

    public class AzhurFormatException : Exception
    {
        public AzhurFormatException() { }

        public AzhurFormatException(string message) : base(message) { }
    }
}
