using Invim.Restxcel.Helpers;
using Invim.Restxcel.Extensions;
using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;

namespace Invim.Restxcel.Models
{
    public class RestxcelDocument : IDisposable
    {
        MemoryStream _str;
        MemoryStream _templateStr;
        ExcelPackage _package;
        RestxcelWorksheetCollection _worksheets;
        private readonly string _id;
        public string Id { get => _id; }
        public Dictionary<string, string> Worksheets { get => _worksheets?.GetListOfWorksheetsWithName(); }

        public RestxcelDocument(RestxcelTemplate template = null)
        {
            _str = new MemoryStream();
            if (template != null)
            {
                _templateStr = new MemoryStream(template.GetData());
                _package = new ExcelPackage(_str, _templateStr);
            }
            else
            {
                _package = new ExcelPackage(_str);
            }
            _worksheets = new RestxcelWorksheetCollection(_package);
            _id = Guid.NewGuid().ToString();
        }

        public void Dispose() => _package?.Dispose();

        public ExcelWorksheet CreateWorksheet(string name, out string id)
        {
            return _worksheets.CreateWorksheet(name, out id);
        }

        public void SetCellValue(string worksheetId, int col, int row, string value)
        {
            if (_worksheets[worksheetId] == null)
            {
                return;
            }
            _worksheets[worksheetId].Cells[row, col].Value = value;
        }

        public void Save() => _package.Save();

        public MemoryStream SaveToStream()
        {
            Save();
            return _str;
        }

        public void SaveToFile(string filePath)
        {
            Save();
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _str.WriteTo(file);
            }
        }

        public ExcelRange GetCell(string worksheetId, int col, int row)
        {
            return _worksheets[worksheetId].Cells[row, col];
        }

        public ExcelRange GetCellRange(string worksheetId, int fromCol, int toCol, int fromRow, int toRow)
        {
            return _worksheets[worksheetId].Cells[fromRow, fromCol, toRow, toCol];
        }

        public void SetCell(string worksheetId, RestxcelCell value)
        {
            var r = _worksheets[worksheetId].Cells[value.Row.Value, value.Col.Value];
            RangeSetCell(r, value);
        }

        public void SetCellRange(string worksheetId, int fromCol, int toCol, int fromRow, int toRow, RestxcelCell value)
        {
            var r = _worksheets[worksheetId].Cells[fromRow, fromCol, toRow, toCol];
            RangeSetCell(r, value);
        }

        private void RangeSetCell(ExcelRange r, RestxcelCell value)
        {
            if (value.Value != null)
            {
                r.Value = value.Value;
            }
            if (!string.IsNullOrEmpty(value.FillColorRgb))
            {
                TryCatchHelper.TryCatchIgnore(() =>
                {
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(value.FillColorRgb.ToColor());
                });
            }
            if (!string.IsNullOrEmpty(value.FontColorRgb))
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.Font.Color.SetColor(value.FontColorRgb.ToColor()));
            }
            if (value.FontSize.HasValue)
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.Font.Size = value.FontSize.Value);
            }
            if (!string.IsNullOrEmpty(value.Font))
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.Font.Name = value.Font);
            }
            if (value.Bold.HasValue)
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.Font.Bold = value.Bold.Value);
            }
            if (value.TextRotation.HasValue)
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.TextRotation = value.TextRotation.Value);
            }
            if (value.ShrinkToFit.HasValue)
            {
                TryCatchHelper.TryCatchIgnore(() => r.Style.ShrinkToFit = value.ShrinkToFit.Value);
            }
        }
    }
}
