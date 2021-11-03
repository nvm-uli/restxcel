using Invim.Restxcel.Models;
using OfficeOpenXml;

namespace Invim.Restxcel.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        public static RestxcelWorksheet ToRestxcelWorksheet(this ExcelWorksheet ws, string id) =>
            new()
            {
                Id = id,
                Name = ws.Name
            };

    }
}
