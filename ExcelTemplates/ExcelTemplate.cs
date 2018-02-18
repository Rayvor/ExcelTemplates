using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelTemplates
{
    public sealed class ExcelTemplate
    {
        private readonly ExcelPackage _xlPackage;

        public ExcelTemplate(string file)
        {
            FileInfo xlTemplate = new FileInfo(file);
            _xlPackage = new ExcelPackage(xlTemplate);
        }

        public byte[] GetAsByteArray()
        {
            return _xlPackage.GetAsByteArray();
        }

        public void SetField(string field, string value)
        {
            foreach (var ws in _xlPackage.Workbook.Worksheets)
            {
                if (ws.Names.ContainsKey(field))
                {
                    ws.Names[field].Value = value;
                    ws.Names[field].AutoFitColumns();
                }
            }

            if (_xlPackage.Workbook.Names.ContainsKey(field))
            {
                _xlPackage.Workbook.Names[field].Value = value;
                _xlPackage.Workbook.Names[field].AutoFitColumns();
            }
        }

        public void SetField(string field, IEnumerable<string> values)
        {
            ExcelRangeBase range = null;
            var namesCollection = _xlPackage.Workbook.Names;

            if (namesCollection.ContainsKey(field))
            {
                var startRow = namesCollection[field].Start.Row;
                var wsIndex = namesCollection[field].Worksheet.Index;

                _xlPackage.Workbook.Worksheets[wsIndex].InsertRow(startRow + 1, values.Count());
                range = namesCollection[field].LoadFromCollection(values, true, OfficeOpenXml.Table.TableStyles.Medium16);
            }

            foreach (var ws in _xlPackage.Workbook.Worksheets)
            {
                if (ws.Names.ContainsKey(field))
                {
                    var data = ws.Names[field].LoadFromCollection(values, true, OfficeOpenXml.Table.TableStyles.Medium16);

                    var table = ws.Tables.GetFromRange(data);
                    table.ShowFilter = true;                    
                }

                if (range != null)
                {
                    var t = ws.Tables.GetFromRange(range);
                    t.ShowFilter = true;
                }
            }
        }       
    }
}
