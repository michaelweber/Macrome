using System;

namespace b2xtranslator.OfficeDrawing
{
    public class InvalidRecordException : Exception
    {
        public InvalidRecordException()
            : base() { }

        public InvalidRecordException(string msg)
            : base(msg) { }

        public InvalidRecordException(string msg, Exception innerException)
            : base(msg, innerException) { }
    }

}
