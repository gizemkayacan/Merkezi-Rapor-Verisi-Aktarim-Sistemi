using DynamicExcelImport.API.Controllers;
using System.Collections.Generic;
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public interface IReportManagementService
    {
        List<object> GetReports(string? roleCode);
        (bool IsSuccess, string Message, string? GeneratedTableName) CreateReport(CreateReportRequest request);
        (bool IsSuccess, byte[]? Content, string? ContentType, string? FileName, string? ErrorMessage) DownloadTemplate(int id);
        List<object> GetImportErrors(long batchId);
        (bool IsSuccess, string Message) DeleteReport(long id);
        (bool IsSuccess, string Message) RenameField(long fieldId, RenameFieldDto dto);
        object? GetReportById(long id);
        List<object> GetFieldsByReportId(long reportId);
        (bool IsSuccess, int TotalRecords, List<Dictionary<string, object>> Data, string? ErrorMessage) GetReportData(long reportId, int pageNumber, int pageSize);
        (bool IsSuccess, string Message) AddFieldsToReport(long id, List<CreateFieldRequest> fields);
        (bool IsSuccess, string Message) DeleteField(long fieldId);
    }
}