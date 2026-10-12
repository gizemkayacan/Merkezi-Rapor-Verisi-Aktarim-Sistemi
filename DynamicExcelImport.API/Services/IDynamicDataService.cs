using DynamicExcelImport.API.Controllers;
using System.Collections.Generic;
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public interface IDynamicDataService
    {
        bool IsReportActive(long reportId);
        long CreateImportBatchRecord(long reportId, string fileName, string fileHash, string status);
        void UpdateImportBatchStatus(long batchId, string status, int totalRows, int validRows, int errorRows);
        List<CreateFieldRequest> GetActiveReportFields(long reportId);
        void LogErrorDetails(long batchId, List<(int RowNumber, string ExcelColumnName, string FieldName, string RawValue, string ErrorMessage)> errorLog);
        bool ExecuteFullRefreshAndInsert(long reportId, long batchId, List<CreateFieldRequest> fields, List<Dictionary<string, object>> parsedRows);
    }
}