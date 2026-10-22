IF COL_LENGTH(N'dbo.Entries', N'TabroomPlace') IS NULL
BEGIN
    ALTER TABLE dbo.Entries ADD TabroomPlace int NULL;
END;

IF COL_LENGTH(N'dbo.Entries', N'PlaceNote') IS NULL
BEGIN
    ALTER TABLE dbo.Entries ADD PlaceNote nvarchar(400) NULL;
END;

EXEC(N'
UPDATE dbo.Entries
SET TabroomPlace = Place
WHERE TabroomPlace IS NULL AND TabroomEntryId IS NOT NULL
');
