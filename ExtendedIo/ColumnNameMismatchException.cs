using System;

namespace ExtendedIo
{
    public class ColumnNameMismatchException : Exception
    {
        public ColumnNameMismatchException()
        {
        }
        public ColumnNameMismatchException(string message)
            : base(message)
        {
        }
        public ColumnNameMismatchException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
