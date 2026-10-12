namespace DynamicExcelImport.UI.Models
{
    public class UpperGroupDto
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public bool IsActive { get; set; } = true;
        public string Description { get; set; }
    }

    public class GroupDto
    {
        public long Id { get; set; }
        public long UpperGroupId { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public bool IsActive { get; set; } = true;
    }

    public class ReportDto
    {
        public long Id { get; set; }
        public long GroupId { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public string PhysicalTableName { get; set; }
        public bool IsActive { get; set; } = true;
        public string RoleCode { get; set; }
    }

    // API'den dönecek hata şablonunu karşılayan akıllı sınıf
    public class ApiErrorResponse
    {
        public long BatchId { get; set; }
        public string Message { get; set; }
        public List<string> Errors { get; set; }
    }
}
// Veritabanından gelen alanları okumak için
public class FieldDto
{
    public long Id { get; set; }
    public long ReportId { get; set; }
    public string Name { get; set; } // Ekranda görünen ad
    public string ExcelColumnName { get; set; } // Excel'deki başlık
    public string Code { get; set; } // SQL'deki kolon adı
    public string DataType { get; set; }
    public int? MaxLength { get; set; }
    public bool IsRequired { get; set; }
    public int OrderNo { get; set; }
}

// Yeni alan eklemek için kullanacağımız form şablonu
public class CreateFieldRequest
{
    public string DisplayName { get; set; }
    public string ExcelColumnName { get; set; }
    public string SqlColumnName { get; set; }
    public string DataType { get; set; }
    public int? MaxLength { get; set; }
    public bool IsRequired { get; set; }
    public int SortOrder { get; set; }
}