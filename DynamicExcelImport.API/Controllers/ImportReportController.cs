using DynamicExcelImport.API.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace DynamicExcelImport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportReportController : ControllerBase
    {
        private readonly ILogger<ImportReportController> _logger;
        private readonly IExcelProcessingService _excelService;
        private readonly IDynamicDataService _dataService;
        private readonly IReportManagementService _reportService;

        public ImportReportController(
            ILogger<ImportReportController> logger,
            IExcelProcessingService excelService,
            IDynamicDataService dataService,
            IReportManagementService reportService)
        {
            _logger = logger;
            _excelService = excelService;
            _dataService = dataService;
            _reportService = reportService;
        }

        [HttpGet]
        public IActionResult GetReports([FromQuery] string? roleCode)
        {
            return Ok(_reportService.GetReports(roleCode));
        }

        [HttpPost]
        public IActionResult CreateReport([FromBody] CreateReportRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Name) || string.IsNullOrWhiteSpace(request.Code))
                return BadRequest("Rapor adı ve kodu boş olamaz!");
            if (request.GroupId <= 0) return BadRequest("Geçerli bir Grup ID'si girmelisiniz!");

            var result = _reportService.CreateReport(request);
            if (result.IsSuccess)
                return Ok(new { Message = result.Message, GeneratedTableName = result.GeneratedTableName });

            return StatusCode(500, result.Message);
        }

        [HttpGet("{id}/template")]
        public IActionResult DownloadTemplate(int id)
        {
            var result = _reportService.DownloadTemplate(id);
            if (!result.IsSuccess) return BadRequest(result.ErrorMessage);

            return File(result.Content, result.ContentType, result.FileName);
        }

        [HttpPost("{id}/imports")]
        public IActionResult ImportExcel(long id, IFormFile file, [FromQuery] int previewCount = 10)
        {
            if (file == null || file.Length == 0) return BadRequest("Lütfen yüklenecek bir Excel dosyası seçin.");
            if (Path.GetExtension(file.FileName).ToLowerInvariant() != ".xlsx") return BadRequest("Sadece .xlsx uzantılı dosyalar yüklenebilir.");

            if (!_dataService.IsReportActive(id))
                return BadRequest("Bu rapor pasif duruma alınmış veya sistemde bulunamadı.");

            string fileHash = "";
            using (var md5 = MD5.Create())
            {
                using (var stream = file.OpenReadStream())
                {
                    var hashBytes = md5.ComputeHash(stream);
                    fileHash = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }

            long batchId = _dataService.CreateImportBatchRecord(id, file.FileName, fileHash, "STARTED");

            var fields = _dataService.GetActiveReportFields(id);
            if (fields.Count == 0)
            {
                _dataService.UpdateImportBatchStatus(batchId, "FAILED", 0, 0, 0);
                return BadRequest("Bu rapora ait aktif alan tanımı bulunamadı.");
            }

            var (parsedRows, errorLog) = _excelService.ParseAndValidate(file, fields);

            if (errorLog.Any())
            {
                _dataService.LogErrorDetails(batchId, errorLog);
                _dataService.UpdateImportBatchStatus(batchId, "FAILED_VALIDATION", parsedRows.Count + errorLog.Count, 0, errorLog.Count);

                var displayErrors = errorLog.Select(e => $"Satır {e.RowNumber} [{e.ExcelColumnName}]: {e.ErrorMessage}").ToList();

                return BadRequest(new
                {
                    BatchId = batchId,
                    Message = "Excel dosyasında verisel hatalar bulundu. İçeri aktarım durduruldu.",
                    Errors = displayErrors
                });
            }

            try
            {
                _dataService.ExecuteFullRefreshAndInsert(id, batchId, fields, parsedRows);
                _dataService.UpdateImportBatchStatus(batchId, "SUCCESS", parsedRows.Count, parsedRows.Count, 0);

                var previewData = parsedRows.Take(previewCount).ToList();

                return Ok(new
                {
                    Message = $"{parsedRows.Count} satır veri başarıyla kaydedildi.",
                    TotalRows = parsedRows.Count,
                    PreviewData = previewData
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "BatchId: {BatchId} numaralı Excel import işlemi sırasında veritabanı hatası.", batchId);
                _dataService.UpdateImportBatchStatus(batchId, "SYSTEM_ERROR", 0, 0, 0);
                return StatusCode(500, $"SQL Hatası: {ex.InnerException?.Message ?? ex.Message}");
            }
        }

        [HttpGet("imports/{batchId}/errors")]
        public IActionResult GetImportErrors(long batchId)
        {
            var errors = _reportService.GetImportErrors(batchId);
            if (errors.Count == 0)
                return NotFound(new { Message = "Hata kaydı bulunamadı veya işlem hatasız tamamlanmış." });

            return Ok(errors);
        }

        [HttpDelete("{id}")]
        public IActionResult DeleteReport(long id)
        {
            var result = _reportService.DeleteReport(id);
            if (result.IsSuccess) return Ok(new { Message = result.Message });

            _logger.LogError("Rapor (Id: {ReportId}) silinirken sistemsel bir hata meydana geldi: {ErrorMessage}", id, result.Message);
            return StatusCode(500, $"Silme işlemi sırasında hata oluştu: {result.Message}");
        }

        [HttpPut("Fields/{fieldId}/Rename")]
        public IActionResult RenameField(long fieldId, [FromBody] RenameFieldDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.DisplayName))
                return BadRequest("Güncellenecek başlık ismi boş olamaz.");

            var result = _reportService.RenameField(fieldId, dto);
            if (result.IsSuccess) return Ok(new { Message = result.Message });

            return StatusCode(500, result.Message);
        }

        [HttpGet("{id}")]
        public IActionResult GetReportById(long id)
        {
            var report = _reportService.GetReportById(id);
            if (report == null) return NotFound($"ID'si {id} olan rapor bulunamadı.");
            return Ok(report);
        }

        [HttpGet("Fields/{reportId}")]
        public IActionResult GetFieldsByReportId(long reportId)
        {
            try
            {
                return Ok(_reportService.GetFieldsByReportId(reportId));
            }
            catch (Exception)
            {
                return StatusCode(500, "Alanlar listelenirken sistemsel bir hata oluştu.");
            }
        }

        [HttpGet("Data/{reportId}")]
        public IActionResult GetReportData(long reportId, [FromQuery] int pageNumber = 1, [FromQuery] int pageSize = 50)
        {
            var result = _reportService.GetReportData(reportId, pageNumber, pageSize);
            if (!result.IsSuccess)
            {
                _logger.LogError("Rapor verileri (ReportId: {ReportId}) çekilirken hata: {Message}", reportId, result.ErrorMessage);
                return result.ErrorMessage == "Rapor bulunamadı."
                    ? NotFound(result.ErrorMessage)
                    : StatusCode(500, $"Veriler listelenirken hata oluştu: {result.ErrorMessage}");
            }

            return Ok(new
            {
                TotalRecords = result.TotalRecords,
                PageNumber = pageNumber,
                PageSize = pageSize,
                Data = result.Data
            });
        }

        [HttpPost("{id}/fields")]
        public IActionResult AddFieldsToReport(long id, [FromBody] List<CreateFieldRequest> fields)
        {
            if (fields == null || !fields.Any()) return BadRequest("Eklenecek alan bilgisi bulunamadı.");

            var result = _reportService.AddFieldsToReport(id, fields);
            if (result.IsSuccess) return Ok(new { Message = result.Message });

            return StatusCode(500, result.Message);
        }

        [HttpDelete("Fields/{fieldId}")]
        public IActionResult DeleteField(long fieldId)
        {
            var result = _reportService.DeleteField(fieldId);
            if (result.IsSuccess) return Ok(new { Message = result.Message });

            return StatusCode(500, result.Message);
        }

        // --- DTO SINIFLARI ---
        public class CreateReportRequest
        {
            public long GroupId { get; set; }
            public string Name { get; set; }
            public string Code { get; set; }
            public string RoleCode { get; set; }
        }

        public class RenameFieldDto
        {
            public string DisplayName { get; set; }
            public string ExcelColumnName { get; set; }
        }

        public class CreateFieldRequest
        {
            public string DisplayName { get; set; }
            public string ExcelColumnName { get; set; }
            public string SqlColumnName { get; set; }
            public string DataType { get; set; }
            public int? MaxLength { get; set; }
            public int? Precision { get; set; }
            public int? Scale { get; set; }
            public bool IsRequired { get; set; }
            public int SortOrder { get; set; }
        }
    }
}