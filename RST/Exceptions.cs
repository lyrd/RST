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
        public class FileException : Exception
        {
            public FileException() { }
            public FileException(string message) : base(message) { }
            public FileException(string message, Exception ex) : base(message) { }
            protected FileException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }
    }
}
