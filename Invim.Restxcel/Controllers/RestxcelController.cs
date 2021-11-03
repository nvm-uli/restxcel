using Invim.Restxcel.Models;
using Invim.Restxcel.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Threading.Tasks;

namespace Invim.Restxcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class RestxcelController : ControllerBase
    {
        private readonly ILogger<RestxcelController> _logger;
        private readonly RestxcelService _restxcel;

        public RestxcelController(ILogger<RestxcelController> logger, RestxcelService restxcel)
        {
            _logger = logger;
            _restxcel = restxcel;
        }

        /// <summary>
        /// Uploads a xltx file (Excel template file) as a template for later use.
        /// </summary>
        /// <param name="permanent">Whether or not to store the template primary on the file system. If not, the template will only persist in the memory of the server.</param>
        /// <returns>Information about the created template including templateId, which is requred to use the template in a new document.</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpPost("UploadTemplate")]
        //[AcceptVerbs("application/octet-stream")]
        [Produces("application/json")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<RestxcelTemplate> UploadTemplate([FromQuery] bool permanent = false)
        {
            MemoryStream ms = new MemoryStream();
            await Request.Body.CopyToAsync(ms);
            return _restxcel.UploadTemplate(ms.ToArray(), permanent);
        }

        /// <summary>
        /// Creates a new Excel workbook. The workbook will remain in server cache until it is downloaded using the "Download" operation.
        /// </summary>
        /// <param name="templateId">Id of a previously uploaded template (see UploadTemplate).</param>
        /// <returns>Information about the created document including documentId, which is required for later use. If a template was used which contained one or more worksheets, the worksheetIds are returned for later use.</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("NewDocument")]
        [Produces("application/json")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public RestxcelDocument NewDocument(string templateId) =>
            _restxcel.CreateDocument(templateId);

        /// <summary>
        /// Creates a new worksheet in the document with the given id.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="name">Name of the new worksheet</param>
        /// <returns>Information about the created worksheet including worksheetId, which is required for later use.</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("NewWorksheet")]
        [Produces("application/json")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public RestxcelWorksheet NewWorksheet([Required] string documentId, [Required] string name) =>
            _restxcel.CreateWorksheet(documentId, name);

        /// <summary>
        /// Returns information about the values and formatting rules for a specific cell in the document.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <param name="col">Col number (starting with 1 [A])</param>
        /// <param name="row">Row number (starting with 1)</param>
        /// <returns>Information about the cell</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("GetCell")]
        [Produces("application/json")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public RestxcelCell GetCell([Required] string documentId, [Required] string worksheetId, [Required] int col, [Required] int row) =>
            _restxcel.GetCell(documentId, worksheetId, col, row);

        /// <summary>
        /// Returns information about the values and formatting rules for a range of cells in the document
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <param name="fromCol">First col number</param>
        /// <param name="toCol">Last col number</param>
        /// <param name="fromRow">First row number</param>
        /// <param name="toRow">Last row number</param>
        /// <returns>Information about the cell</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("GetCellRange")]
        [Produces("application/json")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public RestxcelCell GetCell([Required] string documentId, string worksheetId, [Required] int fromCol, [Required] int toCol, [Required] int fromRow, [Required] int toRow) =>
            _restxcel.GetCellRange(documentId, worksheetId, fromCol, toCol, fromRow, toRow);

        /// <summary>
        /// Simple GET call to apply a value into a specific cell. Use "POST SetCell" for more advanced use
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <param name="col">Col number (starting with 1 [A])</param>
        /// <param name="row">Row number (starting with 1)</param>
        /// <param name="value">The value to be set</param>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("SetCellValue")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public void SetCellValue([Required] string documentId, [Required] string worksheetId, [Required] int col, [Required] int row, [Required] string value = null) =>
            _restxcel.SetCellValue(documentId, worksheetId, col, row, value);

        /// <summary>
        /// Apply a value and formatting rules to a specific cell.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpPost("SetCell")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public void SetCell([Required][FromQuery] string documentId, [Required][FromQuery] string worksheetId, [Required][FromBody] RestxcelCell value) =>
            _restxcel.SetCell(documentId, worksheetId, value);

        /// <summary>
        /// Apply values and formatting rules to a range of cells. This will apply the provided data to all cells within the range. Use BulkSetCell if you want individual rules and values for the cells.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <param name="fromCol">First col number</param>
        /// <param name="toCol">Last col number</param>
        /// <param name="fromRow">First row number</param>
        /// <param name="toRow">Last row number</param>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpPost("SetCellRange")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public void SetCellRange([Required][FromQuery] string documentId, [Required][FromQuery] string worksheetId, [Required][FromQuery] int fromCol, [Required][FromQuery] int toCol, [Required][FromQuery] int fromRow, [Required][FromQuery] int toRow, [Required][FromBody] RestxcelCell value) =>
            _restxcel.SetCellRange(documentId, worksheetId, fromCol, toCol, fromRow, toRow, value);

        /// <summary>
        /// Apply values and formatting rules to a bulk of cells.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="worksheetId">Id of the workbook</param>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpPost("BulkSetCell")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public void BulkSetCell([Required][FromQuery] string documentId, [Required][FromQuery] string worksheetId, [Required][FromBody] IEnumerable<RestxcelCell> value) =>
            _restxcel.BulkSetCell(documentId, worksheetId, value);

        /// <summary>
        /// Downloads a document.
        /// </summary>
        /// <param name="documentId">Id of the document</param>
        /// <param name="fileName">The desired file name including the file extension (.xlsx recommended). Default value: Document.xlsx</param>
        /// <returns>Byte stream with content type "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"</returns>
        /// <response code="200">Returns document as a byte stream</response>
        /// <response code="400">Processing failed with an error. Errors are send as application/json object with the error message displayed in the "error" member</response> 
        [HttpGet("Download")]
        [Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public IActionResult Download([Required] string documentId, string fileName = null) =>
            File(_restxcel.SaveToStream(documentId).ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName ?? "Document.xlsx");
    }
}
