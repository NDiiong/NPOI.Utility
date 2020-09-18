using System;
using System.Runtime.Serialization;

namespace NPOI.Utility.Excel
{
    [Serializable]
    public class SheetNotFoundException : Exception
    {
        public SheetNotFoundException() : base()
        {
        }

        public SheetNotFoundException(string message) : base(message)
        {
        }

        public SheetNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected SheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}