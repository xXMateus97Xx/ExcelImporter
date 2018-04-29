using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelImporter.Exceptions
{
    public class WorksheetNotFoundException : Exception
    {
        public override string Message => "Worksheet was not found on provided file.";
    }
}
