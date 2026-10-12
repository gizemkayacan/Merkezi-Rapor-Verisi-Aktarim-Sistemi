-- Üst Grup Ýçin Eksikler
ALTER TABLE dbo.ImportUpperGroup ADD IsActive BIT NOT NULL DEFAULT 1;
ALTER TABLE dbo.ImportUpperGroup ADD Description NVARCHAR(MAX) NULL;

-- Alt Grup Ýçin Eksikler
ALTER TABLE dbo.ImportGroup ADD IsActive BIT NOT NULL DEFAULT 1;

-- Rapor Ýçin Eksikler
ALTER TABLE dbo.ImportReport ADD IsActive BIT NOT NULL DEFAULT 1;

-- Alan Listesi Ýçin Eksikler
ALTER TABLE dbo.ImportReportField ADD ExcelColumnName NVARCHAR(255) NULL;
ALTER TABLE dbo.ImportReportField ADD OrderNo INT NOT NULL DEFAULT 0;
ALTER TABLE dbo.ImportReportField ADD MaxLength INT NULL;
ALTER TABLE dbo.ImportReportField ADD Precision INT NULL;
ALTER TABLE dbo.ImportReportField ADD DefaultValue NVARCHAR(MAX) NULL;
ALTER TABLE dbo.ImportReportField ADD IsActive BIT NOT NULL DEFAULT 1;
ALTER TABLE dbo.ImportReportField ADD Description NVARCHAR(MAX) NULL;