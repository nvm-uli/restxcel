using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace Invim.Restxcel.Models
{
    public class RestxcelWorksheetCollection
    {
        private readonly ExcelPackage _package;
        private readonly Dictionary<string, ExcelWorksheet> _worksheets;

        public RestxcelWorksheetCollection(ExcelPackage package)
        {
            _package = package;
            if(_package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }
            _worksheets = new Dictionary<string, ExcelWorksheet>();
            if(_package.Workbook.Worksheets != null)
            {
                foreach(var worksheet in _package.Workbook.Worksheets)
                {
                    string id = Guid.NewGuid().ToString();
                    _worksheets.Add(id, worksheet);
                }
            }
        }

        private ExcelWorksheet FindById(string id) => _worksheets[id];

        public ExcelWorksheet this[string id] => FindById(id);

        public ExcelWorksheet CreateWorksheet(string name, out string id)
        {
            var worksheet = _package.Workbook.Worksheets.Add(name ?? "Unnamed");
            id = Guid.NewGuid().ToString();
            _worksheets.Add(id, worksheet);
            return worksheet;
        }

        public Dictionary<string, string> GetListOfWorksheetsWithName()
        {
            Dictionary<string, string> dict = new();
            if(_worksheets != null)
            {
                foreach(var worksheet in _worksheets)
                {
                    dict.Add(worksheet.Key, worksheet.Value.Name);
                }
            }
            return dict;
        }
    }
}
