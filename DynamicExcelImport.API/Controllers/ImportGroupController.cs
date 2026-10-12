using DynamicExcelImport.API.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace DynamicExcelImport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportGroupController : ControllerBase
    {
        private readonly DatabaseService _dbService;

        public ImportGroupController(DatabaseService dbService)
        {
            _dbService = dbService;
        }

        // 1. HTTP GET - Alt Grupları Listeleme
        [HttpGet] // Bu metodun sadece dışarıdan "Bana veri getir" (GET) isteklerine cevap vereceğini söyler.
        [HttpGet]
        // 1. IActionResult yerine async Task<IActionResult> kullanıyoruz
        public async Task<IActionResult> GetGroups()
        {
            var groups = new List<object>();

            using (var connection = (SqlConnection)_dbService.CreateConnection())
            {
                // 2. Open() yerine OpenAsync() kullanıyoruz
                await connection.OpenAsync();

                string query = "SELECT Id, UpperGroupId, Name, Code FROM dbo.ImportGroup WHERE IsActive = 1";

                using (var command = new SqlCommand(query, (SqlConnection)connection))
                {
                    // 3. ExecuteReader() yerine ExecuteReaderAsync() kullanıyoruz
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        // 4. Read() yerine ReadAsync() kullanıyoruz
                        while (await reader.ReadAsync())
                        {
                            groups.Add(new
                            {
                                Id = reader.GetInt64(0),
                                UpperGroupId = reader.GetInt64(1),
                                Name = reader.GetString(2),
                                Code = reader.GetString(3)
                            });
                        }
                    }
                }
            }
            return Ok(groups);
        }

        // 2. HTTP POST - Sisteme Yeni Alt Grup Ekleme (İlişkisel Veri)
        [HttpPost] // Bu metodun sadece sisteme "Yeni veri ekle" (POST) isteklerinde çalışacağını söyler.
        public IActionResult CreateGroup([FromBody] CreateGroupRequest request) // Gelen JSON verisini en alttaki CreateGroupRequest şablonuna oturtup 'request' adıyla içeri alır.
        {
            // Temel boşluk kontrolü
            if (string.IsNullOrWhiteSpace(request.Name) || string.IsNullOrWhiteSpace(request.Code))
            {
                return BadRequest("Alt grup adı ve kodu boş olamaz!"); // İsim veya kod boşsa işlemi reddet (HTTP 400).
            }

            if (request.UpperGroupId <= 0)
            {
                return BadRequest("Geçerli bir Ana Klasör (UpperGroup) ID'si girmelisiniz!"); // ID 0 veya eksiyse reddet.
            }

            using (var connection = _dbService.CreateConnection())
            {
                connection.Open(); // SQL kapısını açtık.

                // --- GÜVENLİK POLİSİ DEVREDE ---
                // Kullanıcının gönderdiği UpperGroupId gerçekten veritabanında var mı diye kontrol ediyoruz
                string checkQuery = "SELECT COUNT(1) FROM dbo.ImportUpperGroup WHERE Id = @UpperGroupId AND IsActive = 1";
                using (var checkCommand = new SqlCommand(checkQuery, (SqlConnection)connection))
                {
                    checkCommand.Parameters.AddWithValue("@UpperGroupId", request.UpperGroupId); // Güvenlik: ID'yi parametre olarak gönderiyoruz (SQL Injection engeli).

                    // ExecuteScalar tek bir değer (burada sayı) dönen sorgular için kullanılır
                    int upperGroupExists = Convert.ToInt32(checkCommand.ExecuteScalar()); // ExecuteScalar: Sadece tek bir hücre (sayı) dönen sorgularda kullanılır. Bize 0 veya 1 döner.

                    if (upperGroupExists == 0) // Eğer 0 döndüyse, öyle bir ana klasör yok demektir!
                    {
                        // Eğer kayıt yoksa işlemi anında iptal et ve hata dön
                        return NotFound($"ID'si {request.UpperGroupId} olan bir Ana Klasör (Üst Grup) bulunamadı. Lütfen önce ana klasörü oluşturun.");
                    }
                }
                // --- GÜVENLİK POLİSİ ONAY VERDİ, KAYDA GEÇEBİLİRİZ ---

                string insertQuery = @"
                    INSERT INTO dbo.ImportGroup (UpperGroupId, Name, Code, IsActive, CreatedAt) 
                    VALUES (@UpperGroupId, @Name, @Code, 1, SYSUTCDATETIME())"; // SYSUTCDATETIME() o anki saat ve tarihi otomatik atar.

                using (var insertCommand = new SqlCommand(insertQuery, (SqlConnection)connection))
                {
                    // Dışarıdan gelen verileri SQL'e güvenli bir şekilde bağlıyoruz.
                    insertCommand.Parameters.AddWithValue("@UpperGroupId", request.UpperGroupId);
                    insertCommand.Parameters.AddWithValue("@Name", request.Name);
                    insertCommand.Parameters.AddWithValue("@Code", request.Code);

                    int affectedRows = insertCommand.ExecuteNonQuery(); // Kayıt işlemini tetikle. Bize etkilenen (eklenen) satır sayısını döner.

                    if (affectedRows > 0) // Eğer 0'dan büyükse kayıt başarıyla eklenmiştir.
                    {
                        return Ok("Yeni alt grup başarıyla ana klasöre bağlandı ve oluşturuldu."); // Başarı mesajı dön (HTTP 200).
                    }
                    else
                    {
                        return StatusCode(500, "Veritabanına kayıt sırasında bir hata oluştu."); // Beklenmedik bir sunucu hatası (HTTP 500).
                    }
                }
            }
        }
        [HttpDelete("{id}")]
        public IActionResult DeleteGroup(long id)
        {
            try
            {
                using (var connection = _dbService.CreateConnection())
                {
                    connection.Open();

                    // KORUMA: İçinde rapor var mı kontrol et
                    string checkQuery = "SELECT COUNT(*) FROM dbo.ImportReport WHERE GroupId = @Id";
                    using (var checkCommand = new SqlCommand(checkQuery, (SqlConnection)connection))
                    {
                        checkCommand.Parameters.AddWithValue("@Id", id);
                        int reportCount = (int)checkCommand.ExecuteScalar();

                        if (reportCount > 0)
                        {
                            return BadRequest($"Bu alt gruba bağlı {reportCount} adet rapor şablonu var. Lütfen önce o raporları silin.");
                        }
                    }

                    // İçinde rapor yoksa güvenle sil
                    string deleteQuery = "DELETE FROM dbo.ImportGroup WHERE Id = @Id";
                    using (var deleteCommand = new SqlCommand(deleteQuery, (SqlConnection)connection))
                    {
                        deleteCommand.Parameters.AddWithValue("@Id", id);
                        deleteCommand.ExecuteNonQuery();
                    }
                }
                return Ok(new { Message = "Alt Grup başarıyla silindi." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Hata: {ex.Message}");
            }
        }
    }

    // Dışarıdan gelen verinin şablonu (DTO) - Bu sefer 3 parametre istiyoruz
    public class CreateGroupRequest
    {
        public int UpperGroupId { get; set; } // Hangi üst gruba bağlanacak?
        public string Name { get; set; } // Alt grubun adı ne?
        public string Code { get; set; } // Alt grubun kodu ne?
    }
}
