using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelImporter.Exceptions
{
    public class ReadOnlyPropertyException : Exception
    {
        public PropertyInfo Property { get; }

        public ReadOnlyPropertyException(PropertyInfo property)
        {
            Property = property;
        }

        public override string Message => $"It's not possible to write on property ${Property.DeclaringType.FullName}.${Property.Name}";
    }
}
