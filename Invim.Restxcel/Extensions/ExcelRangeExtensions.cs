using Invim.Restxcel.Models;
using OfficeOpenXml;

namespace Invim.Restxcel.Extensions
{
    public static class ExcelRangeExtensions
    {
        public static RestxcelCell ToRestxcelCell(this ExcelRange r) =>
            new()
            {
                Value = r.Value?.ToString(),
                FillColorRgb = r.Style.Fill?.BackgroundColor?.Rgb,
                FontColorRgb = r.Style.Font?.Color?.Rgb,
                FontSize = r.Style.Font.Size,
                Address = r.Address
            };
    }
}
