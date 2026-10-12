-- Bu dosya sistemi test etmek amacýyla tamamen sentetik (mock) verilerle oluþturulmuþtur, gerçek kurum verisi içermez.

CREATE TABLE dbo.rpt_finans_bilanco_musteri_akis
(
    Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ImportBatchId BIGINT NOT NULL,
    RowNumber INT NOT NULL,
    
    MusteriNo NVARCHAR(20) NOT NULL,
    FirmaAdi NVARCHAR(250) NOT NULL,
    Bakiye DECIMAL(18,2) NOT NULL,
    CreatedAt DATETIME2(0) NOT NULL DEFAULT SYSUTCDATETIME()
);
INSERT INTO dbo.ImportReportField 
(ReportId, DisplayName, ExcelColumnName, SqlColumnName, DataType, MaxLength, IsRequired, SortOrder, IsActive)
VALUES 
(17, 'Müþteri No', 'Müþteri No', 'MusteriNo', 'string', 20, 1, 1, 1),
(17, 'Firma Adý', 'Firma Adý', 'FirmaAdi', 'string', 250, 1, 2, 1),
(17, 'Bakiye', 'Bakiye', 'Bakiye', 'decimal', NULL, 1, 3, 1);