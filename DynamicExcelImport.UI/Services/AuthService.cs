namespace DynamicExcelImport.UI.Services
{
    public class AuthService
    {
        // Sistem açıldığında varsayılan olarak Risk Departmanı (1000) yetkisiyle başlar
        public string CurrentRoleCode { get; private set; } = "1000";
        public string CurrentUserName { get; private set; } = "Ahmet Yılmaz (Risk Uzmanı)";

        public event Action? OnRoleChanged;

        public void SetRole(string roleCode, string userName)
        {
            CurrentRoleCode = roleCode;
            CurrentUserName = userName;
            OnRoleChanged?.Invoke(); // Rol değiştiğinde ekranı tetikler
        }
    }
}