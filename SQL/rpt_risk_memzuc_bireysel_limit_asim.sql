--Bu dosya sistemi test etmek amacýyla tamamen sentetik (mock) verilerle oluþturulmuþtur, gerçek kurum verisi içermez.

CREATE TABLE dbo.rpt_risk_memzuc_bireysel_limit_asim
(
    Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ImportBatchId BIGINT NOT NULL,
    RowNumber INT NOT NULL,
  
    MusteriNo NVARCHAR(20) unique NOT NULL,
    AdSoyad NVARCHAR(150) NOT NULL,
    TahsisEdilenLimit DECIMAL(18,2) NOT NULL,
    KullanilanRisk DECIMAL(18,2) NOT NULL,
    AsimTutari DECIMAL(18,2) NOT NULL,
    AsimTarihi DATE NULL,
    CreatedAt DATETIME2(0) NOT NULL DEFAULT SYSUTCDATETIME()
);

-- API'nin Excel'i okuyabilmesi için Metadata (Alan Listesi) kayýtlarý
DECLARE @ReportId BIGINT = 14;

INSERT INTO dbo.ImportReportField 
(ReportId, DisplayName, ExcelColumnName, SqlColumnName, DataType, MaxLength, IsRequired, SortOrder, IsActive)
VALUES 
(@ReportId, 'Müþteri No', 'Müþteri No', 'MusteriNo', 'string', 20, 1, 1, 1),
(@ReportId, 'Ad Soyad', 'Ad Soyad', 'AdSoyad', 'string', 150, 1, 2, 1),
(@ReportId, 'Tahsis Edilen Limit', 'Tahsis Edilen Limit', 'TahsisEdilenLimit', 'decimal', NULL, 1, 3, 1),
(@ReportId, 'Kullanýlan Risk', 'Kullanýlan Risk', 'KullanilanRisk', 'decimal', NULL, 1, 4, 1),
(@ReportId, 'Aþým Tutarý', 'Aþým Tutarý', 'AsimTutari', 'decimal', NULL, 1, 5, 1),
(@ReportId, 'Aþým Tarihi', 'Aþým Tarihi', 'AsimTarihi', 'datetime', NULL, 0, 6, 1);