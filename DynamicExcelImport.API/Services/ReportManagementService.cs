using ClosedXML.Excel;
using DynamicExcelImport.API.Controllers;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public class ReportManagementService : IReportManagementService
    {
        private readonly DatabaseService _dbService;

        public ReportManagementService(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        public List<object> GetReports(string? roleCode)
        {
            var reports = new List<object>();
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string query = "SELECT Id, GroupId, Name, Code, PhysicalTableName, SchemaVersion, RoleCode FROM dbo.ImportReport WHERE IsActive = 1";
                if (!string.IsNullOrWhiteSpace(roleCode)) query += " AND RoleCode = @RoleCode";

                using (var command = new SqlCommand(query, (SqlConnection)connection))
                {
                    if (!string.IsNullOrWhiteSpace(roleCode)) command.Parameters.AddWithValue("@RoleCode", roleCode);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            reports.Add(new
                            {
                                Id = Convert.ToInt64(reader["Id"]),
                                GroupId = Convert.ToInt64(reader["GroupId"]),
                                Name = reader["Name"]?.ToString(),
                                Code = reader["Code"]?.ToString(),
                                PhysicalTableName = reader["PhysicalTableName"] != DBNull.Value ? reader["PhysicalTableName"].ToString() : null,
                                SchemaVersion = Convert.ToInt32(reader["SchemaVersion"]),
                                RoleCode = reader["RoleCode"] != DBNull.Value ? reader["RoleCode"].ToString() : null
                            });
                        }
                    }
                }
            }
            return reports;
        }

        //CreateReport
        public (bool IsSuccess, string Message, string? GeneratedTableName) CreateReport(CreateReportRequest request)
        {
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string checkQuery = @"
                        SELECT g.Code AS GroupCode, ug.Code AS UpperGroupCode 
                        FROM dbo.ImportGroup g
                        INNER JOIN dbo.ImportUpperGroup ug ON g.UpperGroupId = ug.Id
                        WHERE g.Id = @GroupId AND g.IsActive = 1";

                string groupCode = "", upperGroupCode = "";

                using (var checkCommand = new SqlCommand(checkQuery, (SqlConnection)connection))
                {
                    checkCommand.Parameters.AddWithValue("@GroupId", request.GroupId);
                    using (var reader = checkCommand.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            groupCode = reader.GetString(0);
                            upperGroupCode = reader.GetString(1);
                        }
                        else return (false, $"ID'si {request.GroupId} olan aktif bir Alt Grup bulunamadı.", null);
                    }
                }

                string physicalTableName = $"rpt_{ToSafeSqlName(upperGroupCode)}_{ToSafeSqlName(groupCode)}_{ToSafeSqlName(request.Code)}";
                string insertQuery = @"
                        INSERT INTO dbo.ImportReport (GroupId, Name, Code, PhysicalTableName, IsActive, SchemaVersion, CreatedAt, RoleCode) 
                        VALUES (@GroupId, @Name, @Code, @PhysicalTableName, 1, 1, SYSDATETIME(), @RoleCode)";

                using (var insertCommand = new SqlCommand(insertQuery, (SqlConnection)connection))
                {
                    insertCommand.Parameters.AddWithValue("@GroupId", request.GroupId);
                    insertCommand.Parameters.AddWithValue("@Name", request.Name);
                    insertCommand.Parameters.AddWithValue("@Code", request.Code);
                    insertCommand.Parameters.AddWithValue("@PhysicalTableName", physicalTableName);
                    insertCommand.Parameters.AddWithValue("@RoleCode", string.IsNullOrEmpty(request.RoleCode) ? DBNull.Value : request.RoleCode);

                    int affectedRows = insertCommand.ExecuteNonQuery();
                    if (affectedRows > 0) return (true, "Yeni rapor başarıyla oluşturuldu.", physicalTableName);
                    else return (false, "Veritabanına rapor kaydı sırasında bir hata oluştu.", null);
                }
            }
        }

        //DownloadTemplate
        public (bool IsSuccess, byte[]? Content, string? ContentType, string? FileName, string? ErrorMessage) DownloadTemplate(int id)
        {
            var fields = new List<CreateFieldRequest>();
            string reportCode = "";
            int schemaVersion = 1;

            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string getReportQuery = "SELECT Code, SchemaVersion FROM dbo.ImportReport WHERE Id = @ReportId AND IsActive = 1";
                using (var cmd = new SqlCommand(getReportQuery, (SqlConnection)connection))
                {
                    cmd.Parameters.AddWithValue("@ReportId", id);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read()) { reportCode = reader.GetString(0); schemaVersion = reader.GetInt32(1); }
                        else return (false, null, null, null, "Aktif rapor bulunamadı.");
                    }
                }
                string getFieldsQuery = "SELECT ExcelColumnName, IsRequired, DataType, DisplayName FROM dbo.ImportReportField WHERE ReportId = @ReportId AND IsActive = 1 ORDER BY SortOrder ASC";
                using (var cmd = new SqlCommand(getFieldsQuery, (SqlConnection)connection))
                {
                    cmd.Parameters.AddWithValue("@ReportId", id);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            fields.Add(new CreateFieldRequest
                            {
                                ExcelColumnName = reader.GetString(0),
                                IsRequired = reader.GetBoolean(1),
                                DataType = reader.GetString(2),
                                DisplayName = reader.GetString(3)
                            });
                        }
                    }
                }
            }
            if (fields.Count == 0) return (false, null, null, null, "Bu rapora ait tanımlanmış aktif bir alan bulunmuyor.");
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Veri_Girisi");
                for (int i = 0; i < fields.Count; i++)
                {
                    int colIndex = i + 1;
                    var cell = worksheet.Cell(1, colIndex);
                    cell.Value = fields[i].ExcelColumnName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.GetComment().AddText($"Veri Tipi: {fields[i].DataType}\nDurum: {(fields[i].IsRequired ? "ZORUNLU" : "Opsiyonel")}");
                }
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    string fileName = $"{reportCode}_sablon_v{schemaVersion}.xlsx";
                    string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    return (true, stream.ToArray(), contentType, fileName, null);
                }
            }
        }

        public List<object> GetImportErrors(long batchId)
        {
            var errors = new List<object>();
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = "SELECT RowNumber, ExcelColumnName, FieldName, RawValue, ErrorMessage FROM dbo.ImportBatchError WHERE BatchId = @BatchId ORDER BY RowNumber ASC";
                using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                {
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            errors.Add(new
                            {
                                RowNumber = reader.GetInt32(0),
                                ExcelColumnName = reader.IsDBNull(1) ? null : reader.GetString(1),
                                FieldName = reader.IsDBNull(2) ? null : reader.GetString(2),
                                RawValue = reader.IsDBNull(3) ? null : reader.GetString(3),
                                ErrorMessage = reader.GetString(4)
                            });
                        }
                    }
                }
            }
            return errors;
        }

        public (bool IsSuccess, string Message) DeleteReport(long id)
        {
            try
            {
                using (var connection = _dbService.CreateConnection())
                {
                    connection.Open();
                    string physicalTableName = null;
                    using (var selectCommand = new SqlCommand("SELECT PhysicalTableName FROM dbo.ImportReport WHERE Id = @Id", (SqlConnection)connection))
                    {
                        selectCommand.Parameters.AddWithValue("@Id", id);
                        var result = selectCommand.ExecuteScalar();
                        if (result == null) return (false, $"ID'si {id} olan rapor bulunamadı.");
                        physicalTableName = result != DBNull.Value ? result.ToString() : null;
                    }

                    if (!string.IsNullOrEmpty(physicalTableName))
                    {
                        using (var dropCommand = new SqlCommand($"DROP TABLE IF EXISTS {physicalTableName}", (SqlConnection)connection))
                            dropCommand.ExecuteNonQuery();
                    }

                    using (var cmd1 = new SqlCommand("DELETE FROM dbo.ImportBatchError WHERE BatchId IN (SELECT Id FROM dbo.ImportBatch WHERE ReportId = @Id)", (SqlConnection)connection))
                    { cmd1.Parameters.AddWithValue("@Id", id); cmd1.ExecuteNonQuery(); }

                    using (var cmd2 = new SqlCommand("DELETE FROM dbo.ImportBatch WHERE ReportId = @Id", (SqlConnection)connection))
                    { cmd2.Parameters.AddWithValue("@Id", id); cmd2.ExecuteNonQuery(); }

                    using (var cmd3 = new SqlCommand("DELETE FROM dbo.ImportReportField WHERE ReportId = @Id", (SqlConnection)connection))
                    { cmd3.Parameters.AddWithValue("@Id", id); cmd3.ExecuteNonQuery(); }

                    using (var cmd4 = new SqlCommand("DELETE FROM dbo.ImportReport WHERE Id = @Id", (SqlConnection)connection))
                    { cmd4.Parameters.AddWithValue("@Id", id); cmd4.ExecuteNonQuery(); }
                }
                return (true, "Rapor şablonu ve ilişkili tüm veri kayıtları başarıyla silindi.");
            }
            catch (Exception ex) { return (false, ex.Message); }
        }

        public (bool IsSuccess, string Message) RenameField(long fieldId, RenameFieldDto dto)
        {
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string updateQuery = "UPDATE dbo.ImportReportField SET DisplayName = @DisplayName, ExcelColumnName = @ExcelColumnName WHERE Id = @FieldId";
                using (var cmd = new SqlCommand(updateQuery, (SqlConnection)connection))
                {
                    cmd.Parameters.AddWithValue("@FieldId", fieldId);
                    cmd.Parameters.AddWithValue("@DisplayName", dto.DisplayName);
                    cmd.Parameters.AddWithValue("@ExcelColumnName", string.IsNullOrEmpty(dto.ExcelColumnName) ? dto.DisplayName : dto.ExcelColumnName);
                    int affectedRows = cmd.ExecuteNonQuery();
                    if (affectedRows == 0) return (false, "Güncellenmek istenen kolon eşleştirmesi bulunamadı.");
                }
            }
            return (true, "Kolon başlık isimleri başarıyla güncellendi.");
        }

        public object? GetReportById(long id)
        {
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                using (var command = new SqlCommand("SELECT Id, Name, Code, PhysicalTableName FROM dbo.ImportReport WHERE Id = @Id", (SqlConnection)connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new
                            {
                                Id = Convert.ToInt64(reader["Id"]),
                                Name = reader["Name"].ToString(),
                                Code = reader["Code"].ToString(),
                                PhysicalTableName = reader["PhysicalTableName"].ToString()
                            };
                        }
                    }
                }
            }
            return null;
        }

        public List<object> GetFieldsByReportId(long reportId)
        {
            var fields = new List<object>();
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string query = "SELECT Id, ReportId, DisplayName, ExcelColumnName, SqlColumnName, DataType, MaxLength, IsRequired, SortOrder, IsActive FROM dbo.ImportReportField WHERE ReportId = @ReportId";
                using (var command = new SqlCommand(query, (SqlConnection)connection))
                {
                    command.Parameters.AddWithValue("@ReportId", reportId);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            fields.Add(new
                            {
                                Id = Convert.ToInt64(reader["Id"]),
                                ReportId = Convert.ToInt64(reader["ReportId"]),
                                Name = reader["DisplayName"].ToString(),
                                ExcelColumnName = reader["ExcelColumnName"] != DBNull.Value ? reader["ExcelColumnName"].ToString() : null,
                                Code = reader["SqlColumnName"].ToString(),
                                DataType = reader["DataType"].ToString(),
                                MaxLength = reader["MaxLength"] != DBNull.Value ? Convert.ToInt32(reader["MaxLength"]) : (int?)null,
                                Precision = (int?)null,
                                IsRequired = Convert.ToBoolean(reader["IsRequired"]),
                                OrderNo = reader["SortOrder"] != DBNull.Value ? Convert.ToInt32(reader["SortOrder"]) : 0,
                                DefaultValue = (string?)null,
                                IsActive = Convert.ToBoolean(reader["IsActive"]),
                                Description = (string?)null
                            });
                        }
                    }
                }
            }
            return fields;
        }

        public (bool IsSuccess, int TotalRecords, List<Dictionary<string, object>> Data, string? ErrorMessage) GetReportData(long reportId, int pageNumber, int pageSize)
        {
            var dataList = new List<Dictionary<string, object>>();
            int totalRecords = 0;

            try
            {
                using (var connection = _dbService.CreateConnection())
                {
                    connection.Open();
                    string getTableQuery = "SELECT PhysicalTableName FROM dbo.ImportReport WHERE Id = @ReportId";
                    string physicalTableName = "";

                    using (var getTableCmd = new SqlCommand(getTableQuery, (SqlConnection)connection))
                    {
                        getTableCmd.Parameters.AddWithValue("@ReportId", reportId);
                        var result = getTableCmd.ExecuteScalar();
                        if (result == null) return (false, 0, dataList, "Rapor bulunamadı.");
                        physicalTableName = result.ToString().ToLowerInvariant();
                    }

                    using (var countCmd = new SqlCommand($"SELECT COUNT(1) FROM dbo.[{physicalTableName}]", (SqlConnection)connection))
                    {
                        totalRecords = (int)countCmd.ExecuteScalar();
                    }

                    int offset = (pageNumber - 1) * pageSize;
                    string dataQuery = $"SELECT * FROM dbo.[{physicalTableName}] ORDER BY Id OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY";

                    using (var dataCmd = new SqlCommand(dataQuery, (SqlConnection)connection))
                    {
                        dataCmd.Parameters.AddWithValue("@Offset", offset);
                        dataCmd.Parameters.AddWithValue("@PageSize", pageSize);

                        using (var reader = dataCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var row = new Dictionary<string, object>();
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (reader.IsDBNull(i)) row[reader.GetName(i)] = null;
                                    else
                                    {
                                        var value = reader.GetValue(i);
                                        if (value is DateTime dt)
                                        {
                                            if (dt.TimeOfDay == TimeSpan.Zero) row[reader.GetName(i)] = dt.ToString("dd.MM.yyyy");
                                            else row[reader.GetName(i)] = dt.ToString("dd.MM.yyyy HH:mm:ss");
                                        }
                                        else row[reader.GetName(i)] = value;
                                    }
                                }
                                dataList.Add(row);
                            }
                        }
                    }
                }
                return (true, totalRecords, dataList, null);
            }
            catch (Exception ex) { return (false, 0, dataList, ex.Message); }
        }

        public (bool IsSuccess, string Message) AddFieldsToReport(long id, List<CreateFieldRequest> fields)
        {
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                using (var checkCmd = new SqlCommand("SELECT COUNT(1) FROM dbo.ImportReport WHERE Id = @Id", (SqlConnection)connection))
                {
                    checkCmd.Parameters.AddWithValue("@Id", id);
                    if (Convert.ToInt32(checkCmd.ExecuteScalar()) == 0) return (false, "Böyle bir rapor bulunamadı.");
                }

                string insertQuery = @"INSERT INTO dbo.ImportReportField (ReportId, DisplayName, ExcelColumnName, SqlColumnName, DataType, MaxLength, Precision, Scale, IsRequired, SortOrder, IsActive) 
                                       VALUES (@ReportId, @DisplayName, @ExcelColumnName, @SqlColumnName, @DataType, @MaxLength, @Precision, @Scale, @IsRequired, @SortOrder, 1)";

                foreach (var field in fields)
                {
                    using (var insertCmd = new SqlCommand(insertQuery, (SqlConnection)connection))
                    {
                        insertCmd.Parameters.AddWithValue("@ReportId", id);
                        insertCmd.Parameters.AddWithValue("@DisplayName", field.DisplayName ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@ExcelColumnName", field.ExcelColumnName ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@SqlColumnName", field.SqlColumnName ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@DataType", field.DataType ?? (object)DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@MaxLength", field.MaxLength.HasValue ? field.MaxLength.Value : DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@Precision", field.Precision.HasValue ? field.Precision.Value : DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@Scale", field.Scale.HasValue ? field.Scale.Value : DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@IsRequired", field.IsRequired);
                        insertCmd.Parameters.AddWithValue("@SortOrder", field.SortOrder);
                        insertCmd.ExecuteNonQuery();
                    }
                }
            }
            return (true, "Alanlar sisteme başarıyla kaydedildi.");
        }

        public (bool IsSuccess, string Message) DeleteField(long fieldId)
        {
            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                using (var cmd = new SqlCommand("DELETE FROM dbo.ImportReportField WHERE Id = @FieldId", (SqlConnection)connection))
                {
                    cmd.Parameters.AddWithValue("@FieldId", fieldId);
                    int affectedRows = cmd.ExecuteNonQuery();
                    if (affectedRows == 0) return (false, "Silinecek alan bulunamadı.");
                }
            }
            return (true, "Alan şablondan başarıyla silindi.");
        }

        private string ToSafeSqlName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            return input.Trim().ToLowerInvariant()
                .Replace("ı", "i").Replace("ğ", "g").Replace("ü", "u").Replace("ş", "s").Replace("ö", "o").Replace("ç", "c")
                .Replace(" ", "_").Replace("-", "_");
        }
    }
}