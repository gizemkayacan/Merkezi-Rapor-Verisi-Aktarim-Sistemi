using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace DynamicExcelImport.API.Services
{
    public class DatabaseService
    {
        // Adresimizi hafızada tutacağımız özel değişken
        private readonly string _connectionString;

        // Yapıcı Metot (Constructor): Bu sınıf ilk çağrıldığında sistem otomatik olarak appsettings.json'ı okur
        // ve içindeki "DefaultConnection" adresini alıp yukarıdaki değişkene kaydeder.
        public DatabaseService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("appsettings.json içinde 'DefaultConnection' bağlantı adresi bulunamadı!");
        }

        // Diğer kodlarımızın veritabanına bağlanmak istediğinde kullanacağı metod
        public SqlConnection CreateConnection()
        {
            // İndirdiğimiz SqlClient paketi sayesinde bu adresi SQL'in anlayacağı bir anahtara çeviriyoruz
            return new SqlConnection(_connectionString);
        }
    }
}