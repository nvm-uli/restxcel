using Invim.Restxcel.Extensions;
using Invim.Restxcel.Models;
using Invim.Restxcel.Settings;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Invim.Restxcel.Services
{
    public class RestxcelService
    {
        private readonly ILogger<RestxcelService> _logger;
        private readonly RestxcelSettings _settings;
        private readonly RestxcelDocumentCollection _documents;
        private readonly RestxcelTemplateCollection _templates;

        public RestxcelService(ILogger<RestxcelService> logger, IOptions<RestxcelSettings> settings)
        {
            _logger = logger;
            _settings = settings?.Value;
            _documents = new();
            _templates = new(_settings);
        }

        public void SetCellValue(string documentId, string worksheetId, int col, int row, string value) =>
            _documents[documentId].SetCellValue(worksheetId, col, row, value: value);

        public RestxcelDocument CreateDocument(string templateId = null) =>
            _documents.NewDocument(out _, string.IsNullOrEmpty(templateId) ? null : _templates[templateId]);

        public RestxcelWorksheet CreateWorksheet(string documentId, string name) =>
            _documents[documentId].CreateWorksheet(name, out var id).ToRestxcelWorksheet(id);

        public void SaveDocument(string documentId, string filePath) =>
            _documents[documentId].SaveToFile(filePath);

        public MemoryStream SaveToStream(string documentId) =>
            _documents[documentId].SaveToStream();

        public RestxcelCell GetCell(string documentId, string worksheetId, int col, int row) =>
            _documents[documentId].GetCell(worksheetId, col, row).ToRestxcelCell();

        public RestxcelCell GetCellRange(string documentId, string worksheetId, int fromCol, int toCol, int fromRow, int toRow) =>
            _documents[documentId].GetCellRange(worksheetId, fromCol, toCol, fromRow, toRow).ToRestxcelCell();

        public void SetCellRange(string documentId, string worksheetId, int fromCol, int toCol, int fromRow, int toRow, RestxcelCell value) =>
            _documents[documentId].SetCellRange(worksheetId, fromCol, toCol, fromRow, toRow, value);

        public void SetCell(string documentId, string worksheetId, RestxcelCell value) =>
            _documents[documentId].SetCell(worksheetId, value);

        public void BulkSetCell(string documentId, string worksheetId, IEnumerable<RestxcelCell> cells) =>
            cells.ToList().ForEach(c => _documents[documentId].SetCell(worksheetId, c));

        public RestxcelTemplate UploadTemplate(byte[] data, bool permanent) =>
            _templates.NewTemplate(data, permanent);
    }
}
