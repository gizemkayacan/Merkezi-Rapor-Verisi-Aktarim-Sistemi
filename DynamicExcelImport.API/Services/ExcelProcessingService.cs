using DynamicExcelImport.API.Controllers;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public class ExcelProcessingService : IExcelProcessingService
    {
        public (List<Dictionary<string, object>> ParsedRows, List<(int RowNumber, string ExcelColumnName, string FieldName, string RawValue, string ErrorMessage)> ErrorLog) ParseAndValidate(IFormFile file, List<CreateFieldRequest> fields)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var errorLog = new List<(int, string, string, string, string)>();
            var parsedRows = new List<Dictionary<string, object>>();

            using (var stream = file.OpenReadStream())
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                if (!reader.Read())
                {
                    errorLog.Add((1, "Genel", null, null, "Excel dosyası tamamen boş."));
                    return (parsedRows, errorLog);
                }

                var headerIndices = new Dictionary<string, int>();
                var excelHeaders = new List<string>();
                var duplicateHeaders = new List<string>();

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var headerValue = reader.GetValue(i)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(headerValue))
                    {
                        if (headerIndices.ContainsKey(headerValue))
                            duplicateHeaders.Add(headerValue);
                        else
                        {
                            headerIndices[headerValue] = i;
                            excelHeaders.Add(headerValue);
                        }
                    }
                }

                if (duplicateHeaders.Any())
                {
                    errorLog.Add((1, "Başlık Hatası", "", "", $"Aynı Excel başlığı birden fazla kez kullanılamaz: {string.Join(", ", duplicateHeaders.Distinct())}"));
                    return (parsedRows, errorLog);
                }

                var expectedHeaders = fields.Select(f => f.ExcelColumnName).ToList();
                var missingHeaders = expectedHeaders.Except(excelHeaders).ToList();
                var extraHeaders = excelHeaders.Except(expectedHeaders).ToList();

                if (missingHeaders.Any() || extraHeaders.Any())
                {
                    if (missingHeaders.Any())
                        errorLog.Add((1, "Başlık Hatası", "", "", $"Şablonda olması gereken ancak Excel'de bulunmayan başlıklar: {string.Join(", ", missingHeaders)}"));
                    if (extraHeaders.Any())
                        errorLog.Add((1, "Başlık Hatası", "", "", $"Sistemin beklemediği ekstra kolonlar bulundu: {string.Join(", ", extraHeaders)}"));
                    return (parsedRows, errorLog);
                }

                int rowNumber = 2;
                while (reader.Read())
                {
                    var rowData = new Dictionary<string, object>();
                    bool isRowEmpty = true;

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (reader.GetValue(i) != null && !string.IsNullOrWhiteSpace(reader.GetValue(i).ToString()))
                        {
                            isRowEmpty = false; break;
                        }
                    }
                    if (isRowEmpty) { rowNumber++; continue; }

                    foreach (var field in fields)
                    {
                        int colIndex = headerIndices[field.ExcelColumnName];
                        var rawValue = reader.GetValue(colIndex);
                        string stringValue = rawValue?.ToString()?.Trim() ?? "";

                        if (field.IsRequired && string.IsNullOrEmpty(stringValue))
                        {
                            errorLog.Add((rowNumber, field.ExcelColumnName, field.SqlColumnName, stringValue, "Bu alan zorunludur ve boş bırakılamaz."));
                            continue;
                        }

                        if (!string.IsNullOrEmpty(stringValue))
                        {
                            if (field.DataType.ToLower() == "string" && field.MaxLength.HasValue && stringValue.Length > field.MaxLength.Value)
                            {
                                errorLog.Add((rowNumber, field.ExcelColumnName, field.SqlColumnName, stringValue, $"Metin en fazla {field.MaxLength.Value} karakter olmalıdır."));
                                continue;
                            }

                            bool isParsed = false;
                            object parsedValue = null;
                            string expectedFormatError = "";

                            switch (field.DataType.ToLower())
                            {
                                case "tam sayı":
                                case "int":
                                    if (long.TryParse(stringValue, out long intVal)) { parsedValue = intVal; isParsed = true; }
                                    else expectedFormatError = "Tam Sayı";
                                    break;

                                case "ondalık":
                                case "decimal":
                                    string safeDecimalStr = stringValue.Replace(",", ".");
                                    if (decimal.TryParse(safeDecimalStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decVal))
                                    {
                                        int precision = field.Precision ?? 18;
                                        int scale = field.Scale ?? 2;
                                        string[] parts = decVal.ToString(CultureInfo.InvariantCulture).Split('.');
                                        int integerPartLength = parts[0].TrimStart('-').Length;
                                        int fractionalPartLength = parts.Length > 1 ? parts[1].Length : 0;
                                        int maxIntegerPartLength = precision - scale;

                                        if (integerPartLength > maxIntegerPartLength)
                                            errorLog.Add((rowNumber, field.ExcelColumnName, field.SqlColumnName, stringValue, $"Sayı sınırı aşıldı! Tam sayı kısmı en fazla {maxIntegerPartLength} basamak olabilir."));
                                        else if (fractionalPartLength > scale)
                                            errorLog.Add((rowNumber, field.ExcelColumnName, field.SqlColumnName, stringValue, $"Ondalık kısım sınırı aşıldı! Virgülden sonra en fazla {scale} basamak olabilir."));
                                        else { parsedValue = decVal; isParsed = true; }
                                    }
                                    else expectedFormatError = "Ondalık Sayı (Örn: 1500.50)";
                                    break;

                                case "tarih":
                                case "tarihsaat":
                                case "datetime":
                                    if (rawValue is DateTime realDate)
                                    {
                                        parsedValue = realDate;
                                        isParsed = true;
                                    }
                                    else
                                    {
                                        string[] dateFormats = { "dd.MM.yyyy", "d.M.yyyy", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy-MM-dd", "MM/dd/yyyy", "yyyy-MM-ddTHH:mm:ss", "dd.MM.yyyy HH:mm:ss" };

                                        if (DateTime.TryParseExact(stringValue, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtVal))
                                        {
                                            parsedValue = dtVal;
                                            isParsed = true;
                                        }
                                        else if (DateTime.TryParse(stringValue, out DateTime genericDate))
                                        {
                                            parsedValue = genericDate;
                                            isParsed = true;
                                        }
                                        else if (double.TryParse(stringValue, out double oaDate) && oaDate > 0)
                                        {
                                            try { parsedValue = DateTime.FromOADate(oaDate); isParsed = true; }
                                            catch { expectedFormatError = "Tarih"; }
                                        }
                                        else
                                        {
                                            expectedFormatError = "Geçerli Tarih Formatı (Örn: dd.MM.yyyy)";
                                        }
                                    }
                                    break;

                                case "evet/hayır":
                                case "boolean":
                                    string lowerStr = stringValue.ToLowerInvariant();
                                    if (lowerStr == "1" || lowerStr == "true" || lowerStr == "evet" || lowerStr == "e") { parsedValue = true; isParsed = true; }
                                    else if (lowerStr == "0" || lowerStr == "false" || lowerStr == "hayır" || lowerStr == "h") { parsedValue = false; isParsed = true; }
                                    else expectedFormatError = "Evet/Hayır veya 1/0";
                                    break;

                                default:
                                    parsedValue = stringValue;
                                    isParsed = true;
                                    break;
                            }

                            if (!isParsed && !string.IsNullOrEmpty(expectedFormatError))
                            {
                                errorLog.Add((rowNumber, field.ExcelColumnName, field.SqlColumnName, stringValue, $"Geçersiz veri. Beklenen tip: {expectedFormatError}."));
                            }
                            else if (isParsed)
                            {
                                rowData[field.SqlColumnName] = parsedValue;
                            }
                        }
                        else
                        {
                            rowData[field.SqlColumnName] = DBNull.Value;
                        }
                    }
                    parsedRows.Add(rowData);
                    rowNumber++;
                }
            }
            return (parsedRows, errorLog);
        }
    }
}