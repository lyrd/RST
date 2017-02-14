using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace RST
{
    class Exceptions
    {
        [Serializable]
        public class CellNameException : Exception
        {
            public CellNameException() { }
            public CellNameException(string message) : base(message) { }
            public CellNameException(string message, Exception ex) : base(message) { }
            protected CellNameException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }
    }
}
