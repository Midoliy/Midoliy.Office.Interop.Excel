using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace Midoliy.Office.Interop
{
    public class AlreadyExistsException : Exception
    {
        public AlreadyExistsException() 
            : base() { }
        public AlreadyExistsException(string message) 
            : base(message) { }
        public AlreadyExistsException(string message, Exception innerException)
            : base(message, innerException) { }
        protected AlreadyExistsException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
