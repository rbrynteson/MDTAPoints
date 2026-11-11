IF COL_LENGTH(N'dbo.Schools', N'ExcludeFromStats') IS NULL
BEGIN
    ALTER TABLE dbo.Schools ADD ExcludeFromStats bit NOT NULL
        CONSTRAINT DF_Schools_ExcludeFromStats DEFAULT 0;
END;
