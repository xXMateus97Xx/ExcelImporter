using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;

namespace ExcelImporter
{
    internal class MapperConfiguration
    {
        public MapperConfiguration(IList<MapperFieldConfiguration> fields, string worksheetName)
        {
            Fields = new ReadOnlyCollection<MapperFieldConfiguration>(fields);
            WorksheetName = worksheetName;
        }

        public IReadOnlyCollection<MapperFieldConfiguration> Fields { get; }
        public string WorksheetName { get; }
    }

    internal abstract class MapperFieldConfiguration
    {
        internal MapperFieldConfiguration(string excelName, int displayOrder, string sampleData)
        {
            ExcelName = excelName;
            DisplayOrder = displayOrder;
            SampleData = sampleData;
        }

        public abstract Type Type { get; }
        public string ExcelName { get; }
        public int DisplayOrder { get; }
        public string SampleData { get; }
    }

    internal class MapperClassFieldConfiguration : MapperFieldConfiguration
    {
        internal MapperClassFieldConfiguration(PropertyInfo prop, string excelName, int displayOrder, string sampleData)
            : base(excelName, displayOrder, sampleData)
        {
            Property = prop;
        }

        public PropertyInfo Property { get; }

        public override Type Type => Property.PropertyType;
    }

    internal class MapperExtraFieldConfiguration : MapperFieldConfiguration
    {
        internal MapperExtraFieldConfiguration(string systemName, Type type, string excelName, int displayOrder, string sampleData)
            : base(excelName, displayOrder, sampleData)
        {
            SystemName = systemName;
            Type = type;
        }

        public override Type Type { get; }
        public string SystemName { get; }
    }
}
