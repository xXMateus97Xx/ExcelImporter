using ExcelImporter.Exceptions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExcelImporter
{
    public class Importer<T> : IImporter<T> where T : class, new()
    {
        private readonly Type _typeOfT;
        private MapperConfiguration _configuration;

        public Importer()
        {
            _typeOfT = typeof(T);
            _configuration = ImporterMapper.GetConfiguration(_typeOfT);
        }

        public byte[] GetSampleExcelFile()
        {
            var stream = new MemoryStream();
            using (var xlPackage = new ExcelPackage(stream))
            {
                var worksheet = xlPackage.Workbook.Worksheets.Add(_configuration.WorksheetName);

                var properties = _configuration.Fields.OrderBy(x => x.DisplayOrder);

                var column = 1;
                foreach (var property in properties)
                {
                    SetTitleRowStyle(worksheet.Cells[1, column].Style);
                    worksheet.Cells[1, column].Value = property.ExcelName;
                    worksheet.Cells[2, column].Value = property.SampleData;
                    column++;
                }

                xlPackage.Save();
                return stream.ToArray();
            }
        }

        private void SetTitleRowStyle(ExcelStyle style)
        {
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            style.Font.Bold = true;
        }

        public IList<ImportedItem<T>> Import(Stream excel)
        {
            if (excel == null)
                throw new ArgumentNullException(nameof(excel));

            var result = new List<ImportedItem<T>>();

            using (var xlPackage = new ExcelPackage(excel))
            {
                var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                    throw new WorksheetNotFoundException();

                var row = 2;

                while (!AreAllColumnsEmpty(worksheet, row, _configuration.Fields.Count))
                {
                    var value = new T();
                    var extraFields = new Dictionary<string, object>();
                    var errors = new List<string>();

                    foreach (var item in _configuration.Fields)
                    {
                        var column = GetColumnIndex(worksheet, item.ExcelName);
                        var excelValue = worksheet.Cells[row, column].Value;
                        object rawValue;
                        Type type;

                        try
                        {
                            var isNullable = false;
                            var nullType = Nullable.GetUnderlyingType(item.Type);
                            if (nullType != null)
                            {
                                isNullable = true;
                                type = nullType;
                            }
                            else
                            {
                                type = item.Type;
                            }

                            if (isNullable)
                            {
                                if (excelValue == null)
                                    rawValue = null;
                                else
                                    rawValue = Convert.ChangeType(excelValue, nullType, Thread.CurrentThread.CurrentCulture);
                            }
                            else if (type == typeof(DateTime))
                            {
                                if (double.TryParse(excelValue.ToString(), out var timestamp))
                                    rawValue = DateTime.FromOADate(timestamp);
                                else
                                    rawValue = Convert.ChangeType(excelValue, type, Thread.CurrentThread.CurrentCulture);
                            }
                            else
                            {
                                rawValue = Convert.ChangeType(excelValue, type, Thread.CurrentThread.CurrentCulture);
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add(ex.Message);
                            rawValue = item.Type.GetDefaultValue();
                        }

                        if (item is MapperClassFieldConfiguration classFieldConfiguration)
                            classFieldConfiguration.Property.SetValue(value, rawValue);
                        else if (item is MapperExtraFieldConfiguration extraFieldConfiguration)
                            extraFields.Add(extraFieldConfiguration.SystemName, rawValue);
                    }

                    result.Add(new ImportedItem<T>(value, extraFields, errors));
                    row++;
                }
            }

            return result;
        }

        private bool AreAllColumnsEmpty(ExcelWorksheet worksheet, int row, int lastColumn)
        {
            for (var i = 1; i <= lastColumn; i++)
            {
                if (!string.IsNullOrEmpty(worksheet.Cells[row, i].Value?.ToString()))
                    return false;
            }

            return true;
        }

        private int GetColumnIndex(ExcelWorksheet worksheet, string label)
        {
            var column = 0;
            var value = "dummy";

            while (!string.IsNullOrEmpty(value))
            {
                var cell = worksheet.Cells[1, ++column];
                value = cell.Value.ToString();
                if (value == label)
                    return column;
            }

            return 0;
        }

        public IList<ImportedItem<T>> Import(byte[] excel)
        {
            using (var stream = new MemoryStream(excel))
                return Import(stream);
        }
    }
}
