using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Linq;
using System.Reflection;
using ExcelImporter.Exceptions;

namespace ExcelImporter
{
    public class ImporterMapper
    {
        private static readonly ConcurrentDictionary<Type, MapperConfiguration> _configurations;

        static ImporterMapper()
        {
            _configurations = new ConcurrentDictionary<Type, MapperConfiguration>();
        }

        internal static MapperConfiguration GetConfiguration(Type type)
        {
            if (_configurations.TryGetValue(type, out var conf))
                return conf;

            throw new ConfigurationNotFoundException(type);
        }

        internal static void AddMapper(Type type, MapperConfiguration configuration)
        {
            _configurations.TryAdd(type, configuration);
        }
    }

    public class ImporterMapper<T> : ImporterMapper where T : class, new()
    {
        private IList<MapperFieldConfiguration> _fields;
        private IList<PropertyInfo> _ignoredFields;
        private string _worksheetName;

        private ImporterMapper()
        {
            _fields = new List<MapperFieldConfiguration>();
            _ignoredFields = new List<PropertyInfo>();
            _worksheetName = typeof(T).Name;
        }

        public static ImporterMapper<T> Create()
        {
            return new ImporterMapper<T>();
        }

        public ImporterMapper<T> AddField<TProp>(Expression<Func<T, TProp>> prop, string excelName, int displayOrder, string sampleData)
        {
            var propInfo = GetPropInfo(prop);
            if (!propInfo.CanWrite)
                throw new ReadOnlyPropertyException(propInfo);
            _fields.Add(new MapperClassFieldConfiguration(propInfo, excelName, displayOrder, sampleData));

            return this;
        }

        public ImporterMapper<T> AddExtraField(string systemName, Type type, string excelName, int displayOrder, string sampleData)
        {
            _fields.Add(new MapperExtraFieldConfiguration(systemName, type, excelName, displayOrder, sampleData));

            return this;
        }

        public ImporterMapper<T> Ignore<TProp>(Expression<Func<T, TProp>> prop)
        {
            var propInfo = GetPropInfo(prop);
            _ignoredFields.Add(propInfo);

            return this;
        }

        public ImporterMapper<T> SetWorksheetName(string name)
        {
            if (!string.IsNullOrEmpty(name))
                _worksheetName = name;
            return this;
        }

        public void Map()
        {
            var typeOfT = typeof(T);
            var classProperties = _fields.Where(x => x is MapperClassFieldConfiguration)
                .Select(x => x as MapperClassFieldConfiguration).ToList();

            foreach (var prop in typeOfT.GetProperties().Where(x => x.CanWrite))
            {
                if (!classProperties.Any(x => x.Property == prop) && !_ignoredFields.Contains(prop))
                {
                    _fields.Add(new MapperClassFieldConfiguration(prop, prop.Name, 9999,
                                                                  prop.PropertyType.GetDefaultValue()?.ToString()));
                }
            }

            AddMapper(typeOfT, new MapperConfiguration(_fields, _worksheetName));
        }

        private PropertyInfo GetPropInfo<TProp>(Expression<Func<T, TProp>> prop)
        {
            if (prop.Body is MemberExpression && (prop.Body as MemberExpression).Member is PropertyInfo)
                return (prop.Body as MemberExpression).Member as PropertyInfo;

            throw new Exception();
        }
    }
}
