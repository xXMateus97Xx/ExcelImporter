using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelImporter
{
    public static class ReflectionExtensions
    {
        public static object GetDefaultValue(this Type type)
        {
            if (type.IsValueType)
                return Activator.CreateInstance(type);

            return null;
        }
    }
}
