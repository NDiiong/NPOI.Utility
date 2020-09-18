using System;
using System.Runtime.Serialization;

namespace NPOI.Utility.Excel
{
    [Serializable]
    public class CellNotFoundException : Exception
    {
        public CellNotFoundException() : base()
        {
        }

        public CellNotFoundException(string message) : base(message)
        {
        }

        public CellNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected CellNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}