using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelImporter.Exceptions
{
    public class ExpressionCastException : Exception
    {
        private string _argName;

        public ExpressionCastException(string argName)
        {
            _argName = argName;
        }

        public override string Message => $"{_argName} is not a Property";
    }
}
