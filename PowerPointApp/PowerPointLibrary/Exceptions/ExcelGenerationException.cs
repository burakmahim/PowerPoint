using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Exceptions
{
    public class ExcelGenerationException : Exception
    {
        public ExcelGenerationException(string message, Exception? innerException = null)
            : base(message, innerException)
        {
        }
    }
}