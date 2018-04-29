using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelImporter.Exceptions
{
    public class ConfigurationNotFoundException : Exception
    {
        public Type Type { get; }

        public ConfigurationNotFoundException(Type type)
        {
            Type = type;
        }

        public override string Message => $"Configuration not found for {Type.FullName}";
    }
}
