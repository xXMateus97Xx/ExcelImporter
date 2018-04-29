using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelImporter
{
    public interface IImporter<T> where T : class, new()
    {
        IList<ImportedItem<T>> Import(Stream excel);
        IList<ImportedItem<T>> Import(byte[] excel);
        byte[] GetSampleExcelFile();
    }
}
