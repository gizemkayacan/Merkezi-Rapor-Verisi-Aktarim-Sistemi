using DynamicExcelImport.API.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

// NOT: Kendi projende namespace "DynamicExcellImport.API.Controllers" ise burayı ona göre düzeltmeyi unutma
namespace DynamicExcelImport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportUpperGroupController : ControllerBase
    {
        private readonly DatabaseService _dbService;

        public ImportUpperGroupController(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        // 1. HTTP GET - Veri Listeleme
        [HttpGet]
        public IActionResult GetUpperGroups()
        {
            var groups = new List<object>();

            using (var connection = _dbService.CreateConnection())
            {
                connection.Open();
                string query = "SELECT Id, Name, Code FROM dbo.ImportUpperGroup WHERE IsActive = 1";

                using (var command = new SqlCommand(query, (SqlConnection)connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            groups.Add(new
                            {
                                Id = reader.GetInt64(0),
                                Name = reader.GetString(1),
                                Code = reader.GetString(2)
                            });
                        }
                    }
                }
            }
            return Ok(groups);
        }

        [HttpPost]
        public async Task<IActionResult> CreateUpperGroup([FromBody] CreateUpperGroupRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Name) || string.IsNullOrWhiteSpace(request.Code))
                return BadRequest("Grup adı ve kodu boş olamaz!");

            using (var connection = _dbService.CreateConnection())
            {
                await connection.OpenAsync();

                string query = @"
        INSERT INTO dbo.ImportUpperGroup (Name, Code, IsActive, CreatedAt, UpdatedAt) 
        VALUES (@Name, @Code, 1, SYSUTCDATETIME(), SYSUTCDATETIME())";

                using (var command = new SqlCommand(query, (SqlConnection)connection))
                {
                    command.Parameters.AddWithValue("@Name", request.Name);
                    command.Parameters.AddWithValue("@Code", request.Code);

                    int affectedRows = await command.ExecuteNonQueryAsync();

                    if (affectedRows > 0)
                        return Ok("Yeni üst grup başarıyla oluşturuldu.");
                    else
                        return StatusCode(500, "Veritabanına kayıt sırasında bir hata oluştu.");
                }
            }
        }

        [HttpDelete("{id}")]
        public IActionResult DeleteUpperGroup(long id)
        {
            try
            {
                using (var connection = _dbService.CreateConnection())
                {
                    connection.Open();

                    // KORUMA: İçinde alt grup var mı kontrol et
                    string checkQuery = "SELECT COUNT(*) FROM dbo.ImportGroup WHERE UpperGroupId = @Id";
                    using (var checkCommand = new SqlCommand(checkQuery, (SqlConnection)connection))
                    {
                        checkCommand.Parameters.AddWithValue("@Id", id);
                        int groupCount = (int)checkCommand.ExecuteScalar();

                        if (groupCount > 0)
                        {
                            return BadRequest($"Bu üst gruba bağlı {groupCount} adet alt grup var. Lütfen önce o alt grupları silin.");
                        }
                    }

                    // İçinde alt grup yoksa güvenle sil
                    string deleteQuery = "DELETE FROM dbo.ImportUpperGroup WHERE Id = @Id";
                    using (var deleteCommand = new SqlCommand(deleteQuery, (SqlConnection)connection))
                    {
                        deleteCommand.Parameters.AddWithValue("@Id", id);
                        deleteCommand.ExecuteNonQuery();
                    }
                }
                return Ok(new { Message = "Üst Grup başarıyla silindi." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Hata: {ex.Message}");
            }
        }
    }

    // Dışarıdan gelen verinin şablonu (DTO)
    public class CreateUpperGroupRequest
    {
        public string Name { get; set; }
        public string Code { get; set; }
        public string? Description { get; set; }
    }
}