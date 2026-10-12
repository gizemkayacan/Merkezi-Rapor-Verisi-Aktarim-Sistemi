using DynamicExcelImport.API.Controllers;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
// CreateFieldRequest sınıfını Controller'dan okuyabilmesi için:
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public interface IExcelProcessingService
    {
        // Tuple yapısıyla hem verileri hem de hata loglarını dönecek sözleşmemiz
        (List<Dictionary<string, object>> ParsedRows, List<(int RowNumber, string ExcelColumnName, string FieldName, string RawValue, string ErrorMessage)> ErrorLog) ParseAndValidate(IFormFile file, List<CreateFieldRequest> fields);
    }
}