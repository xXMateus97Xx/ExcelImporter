using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace ExcelImporter
{
    public class ImportedItem
    {
        public ImportedItem(object value, IDictionary<string, object> extraFields, IList<string> errors)
        {
            ExtraFields = new ReadOnlyDictionary<string, object>(extraFields);
            Value = value;
            Errors = errors;
        }

        public object Value { get; }

        public IReadOnlyDictionary<string, object> ExtraFields { get; }

        public IList<string> Errors { get; }

        public bool HasError => Errors.Any();
    }

    public class ImportedItem<T> : ImportedItem where T : class
    {
        public ImportedItem(T value, IDictionary<string, object> extraFields, IList<string> errors) 
            : base(value, extraFields, errors)
        {

        }

        public new T Value => (T)base.Value;
    }
}
