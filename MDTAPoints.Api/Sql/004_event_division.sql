IF COL_LENGTH(N'dbo.Events', N'Division') IS NULL
BEGIN
    ALTER TABLE dbo.Events ADD Division nvarchar(20) NOT NULL
        CONSTRAINT DF_Events_Division DEFAULT N'Varsity';
END;
