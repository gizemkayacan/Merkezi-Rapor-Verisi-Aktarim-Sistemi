using DynamicExcelImport.API.Controllers;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using static DynamicExcelImport.API.Controllers.ImportReportController;

namespace DynamicExcelImport.API.Services
{
    public class DynamicDataService : IDynamicDataService
    {
        private readonly DatabaseService _dbService;
        public DynamicDataService(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        public bool IsReportActive(long reportId)
        {
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = "SELECT IsActive FROM dbo.ImportReport WHERE Id = @Id";
                using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                {
                    cmd.Parameters.AddWithValue("@Id", reportId);
                    var result = cmd.ExecuteScalar();
                    return result != null && Convert.ToBoolean(result);
                }
            }
        }

        public long CreateImportBatchRecord(long reportId, string fileName, string fileHash, string status)
        {
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = @"
            INSERT INTO dbo.ImportBatch (ReportId, FileName, FileHash, Status, StartedAt) 
            OUTPUT INSERTED.Id 
            VALUES (@ReportId, @FileName, @FileHash, @Status, SYSUTCDATETIME())";

                using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                {
                    cmd.Parameters.AddWithValue("@ReportId", reportId);
                    cmd.Parameters.AddWithValue("@FileName", fileName);
                    cmd.Parameters.AddWithValue("@FileHash", string.IsNullOrEmpty(fileHash) ? DBNull.Value : fileHash);
                    cmd.Parameters.AddWithValue("@Status", status);

                    return (long)cmd.ExecuteScalar();
                }
            }
        }

        public void UpdateImportBatchStatus(long batchId, string status, int totalRows, int validRows, int errorRows)
        {
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = "UPDATE dbo.ImportBatch SET Status = @Status, TotalRows = @TotalRows, ValidRows = @ValidRows, ErrorRows = @ErrorRows, FinishedAt = SYSUTCDATETIME() WHERE Id = @Id";
                using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                {
                    cmd.Parameters.AddWithValue("@Id", batchId);
                    cmd.Parameters.AddWithValue("@Status", status);
                    cmd.Parameters.AddWithValue("@TotalRows", totalRows);
                    cmd.Parameters.AddWithValue("@ValidRows", validRows);
                    cmd.Parameters.AddWithValue("@ErrorRows", errorRows);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public List<CreateFieldRequest> GetActiveReportFields(long reportId)
        {
            var fields = new List<CreateFieldRequest>();
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = @"
                    SELECT ExcelColumnName, IsRequired, DataType, DisplayName, MaxLength, SqlColumnName 
                    FROM dbo.ImportReportField 
                    WHERE ReportId = @ReportId AND IsActive = 1 
                    ORDER BY SortOrder ASC";

                using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                {
                    cmd.Parameters.AddWithValue("@ReportId", reportId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            fields.Add(new CreateFieldRequest
                            {
                                ExcelColumnName = reader.GetString(0),
                                IsRequired = reader.GetBoolean(1),
                                DataType = reader.GetString(2),
                                DisplayName = reader.GetString(3),
                                MaxLength = reader.IsDBNull(4) ? (int?)null : reader.GetInt32(4),
                                SqlColumnName = reader.GetString(5)
                            });
                        }
                    }
                }
            }
            return fields;
        }

        public void LogErrorDetails(long batchId, List<(int RowNumber, string ExcelColumnName, string FieldName, string RawValue, string ErrorMessage)> errorLog)
        {
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                string query = "INSERT INTO dbo.ImportBatchError (BatchId, RowNumber, ExcelColumnName, FieldName, RawValue, ErrorMessage) VALUES (@BatchId, @RowNumber, @ExcelColumnName, @FieldName, @RawValue, @ErrorMessage)";
                foreach (var err in errorLog)
                {
                    using (var cmd = new SqlCommand(query, (SqlConnection)conn))
                    {
                        cmd.Parameters.AddWithValue("@BatchId", batchId);
                        cmd.Parameters.AddWithValue("@RowNumber", err.RowNumber);
                        cmd.Parameters.AddWithValue("@ExcelColumnName", string.IsNullOrEmpty(err.ExcelColumnName) ? DBNull.Value : err.ExcelColumnName);
                        cmd.Parameters.AddWithValue("@FieldName", string.IsNullOrEmpty(err.FieldName) ? DBNull.Value : err.FieldName);
                        cmd.Parameters.AddWithValue("@RawValue", string.IsNullOrEmpty(err.RawValue) ? DBNull.Value : err.RawValue);
                        cmd.Parameters.AddWithValue("@ErrorMessage", err.ErrorMessage);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        //ExecuteFullRefreshAndInsert
        public bool ExecuteFullRefreshAndInsert(long reportId, long batchId, List<CreateFieldRequest> fields, List<Dictionary<string, object>> parsedRows)
        {
            using (var conn = _dbService.CreateConnection())
            {
                conn.Open();
                using (var transaction = ((SqlConnection)conn).BeginTransaction())
                {
                    try
                    {
                        string tableName = "";
                        using (var cmd = new SqlCommand("SELECT PhysicalTableName FROM dbo.ImportReport WHERE Id = @ReportId", (SqlConnection)conn, transaction))
                        {
                            cmd.Parameters.AddWithValue("@ReportId", reportId);
                            tableName = cmd.ExecuteScalar()?.ToString();
                        }

                        if (string.IsNullOrEmpty(tableName)) throw new Exception("Tablo adı bulunamadı.");

                        using (var delCmd = new SqlCommand($"DELETE FROM dbo.{tableName}", (SqlConnection)conn, transaction))
                        {
                            delCmd.ExecuteNonQuery();
                        }
                        DataTable dataTable = new DataTable();
                        dataTable.Columns.Add("ImportBatchId", typeof(long));
                        dataTable.Columns.Add("RowNumber", typeof(int));

                        foreach (var field in fields)
                        {
                            dataTable.Columns.Add(field.SqlColumnName, typeof(object));
                        }
                        int currentRowNum = 2;
                        foreach (var row in parsedRows)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            dataRow["ImportBatchId"] = batchId;
                            dataRow["RowNumber"] = currentRowNum;

                            foreach (var field in fields)
                            {
                                dataRow[field.SqlColumnName] = row.ContainsKey(field.SqlColumnName) && row[field.SqlColumnName] != null
                                    ? row[field.SqlColumnName]
                                    : DBNull.Value;
                            }

                            dataTable.Rows.Add(dataRow);
                            currentRowNum++;
                        }

                        using (var bulkCopy = new SqlBulkCopy((SqlConnection)conn, SqlBulkCopyOptions.Default, transaction))
                        {
                            bulkCopy.DestinationTableName = $"dbo.{tableName}";
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                            }
                            bulkCopy.WriteToServer(dataTable);
                        }

                        transaction.Commit();
                        return true;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}