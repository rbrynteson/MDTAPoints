IF OBJECT_ID(N'dbo.SchemaVersion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SchemaVersion
    (
        Version int NOT NULL CONSTRAINT PK_SchemaVersion PRIMARY KEY,
        AppliedAt datetime2 NOT NULL
    );
END;

IF OBJECT_ID(N'dbo.Schools', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Schools
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Schools PRIMARY KEY,
        TabroomSchoolId int NULL,
        Name nvarchar(300) NOT NULL,
        NormalizedName nvarchar(300) NOT NULL
    );
    CREATE UNIQUE INDEX UX_Schools_Tabroom ON dbo.Schools(TabroomSchoolId) WHERE TabroomSchoolId IS NOT NULL;
END;

IF OBJECT_ID(N'dbo.Students', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Students
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Students PRIMARY KEY,
        SchoolId int NOT NULL CONSTRAINT FK_Students_Schools REFERENCES dbo.Schools(Id),
        DisplayName nvarchar(300) NOT NULL,
        NormalizedName nvarchar(300) NOT NULL,
        Grade int NULL
    );
    CREATE INDEX IX_Students_School_Name ON dbo.Students(SchoolId, NormalizedName);
END;

IF OBJECT_ID(N'dbo.StudentAliases', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.StudentAliases
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_StudentAliases PRIMARY KEY,
        StudentId int NOT NULL CONSTRAINT FK_Aliases_Students REFERENCES dbo.Students(Id),
        NormalizedName nvarchar(300) NOT NULL,
        CONSTRAINT UQ_StudentAliases UNIQUE (StudentId, NormalizedName)
    );
END;

IF OBJECT_ID(N'dbo.StudentQualifications', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.StudentQualifications
    (
        StudentId int NOT NULL CONSTRAINT FK_Qual_Students REFERENCES dbo.Students(Id),
        EventKind nvarchar(40) NOT NULL,
        QualifiedToState bit NOT NULL,
        CONSTRAINT PK_StudentQualifications PRIMARY KEY (StudentId, EventKind)
    );
END;

IF OBJECT_ID(N'dbo.HeadToHead', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.HeadToHead
    (
        WinnerStudentId int NOT NULL CONSTRAINT FK_H2H_Winner REFERENCES dbo.Students(Id),
        LoserStudentId int NOT NULL CONSTRAINT FK_H2H_Loser REFERENCES dbo.Students(Id),
        CONSTRAINT PK_HeadToHead PRIMARY KEY (WinnerStudentId, LoserStudentId)
    );
END;

IF OBJECT_ID(N'dbo.Users', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Users
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Users PRIMARY KEY,
        Username nvarchar(100) NOT NULL CONSTRAINT UQ_Users_Username UNIQUE,
        DisplayName nvarchar(200) NOT NULL,
        PasswordHash nvarchar(500) NOT NULL
    );
END;

IF OBJECT_ID(N'dbo.Tournaments', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tournaments
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Tournaments PRIMARY KEY,
        TabroomTournamentId int NULL,
        Name nvarchar(400) NOT NULL,
        StartDate date NOT NULL,
        StateCode nvarchar(20) NULL,
        InState bit NOT NULL,
        Kind nvarchar(40) NOT NULL,
        Status nvarchar(20) NOT NULL,
        SkippedSummary nvarchar(max) NULL,
        RawJson nvarchar(max) NULL
    );
END;

IF OBJECT_ID(N'dbo.Events', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Events
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Events PRIMARY KEY,
        TournamentId int NOT NULL CONSTRAINT FK_Events_Tournaments REFERENCES dbo.Tournaments(Id),
        TabroomEventId int NULL,
        TabroomResultSetId int NULL,
        Name nvarchar(400) NOT NULL,
        EventKind nvarchar(40) NOT NULL,
        CountedDivision bit NOT NULL,
        SchoolCount int NOT NULL,
        EntryCount int NOT NULL,
        BreaksToOctosOrDeeper bit NOT NULL,
        HasFinalSession bit NOT NULL,
        HasSemis bit NOT NULL,
        IsHouse bit NOT NULL,
        RawJson nvarchar(max) NULL
    );
END;

IF OBJECT_ID(N'dbo.Entries', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Entries
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Entries PRIMARY KEY,
        EventId int NOT NULL CONSTRAINT FK_Entries_Events REFERENCES dbo.Events(Id),
        TabroomEntryId int NULL,
        EntryName nvarchar(400) NOT NULL,
        SchoolId int NOT NULL CONSTRAINT FK_Entries_Schools REFERENCES dbo.Schools(Id),
        Place int NOT NULL,
        IsPrelim bit NOT NULL,
        ReachedElims bit NOT NULL,
        ElimWins int NOT NULL,
        ReachedFinals bit NOT NULL,
        ReachedSemis bit NOT NULL,
        QualifiedToState bit NOT NULL
    );
END;

IF OBJECT_ID(N'dbo.EntryCompetitors', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EntryCompetitors
    (
        Id int IDENTITY(1,1) NOT NULL CONSTRAINT PK_EntryCompetitors PRIMARY KEY,
        EntryId int NOT NULL CONSTRAINT FK_Competitors_Entries REFERENCES dbo.Entries(Id),
        RawName nvarchar(300) NOT NULL,
        NormalizedName nvarchar(300) NOT NULL,
        StudentId int NULL CONSTRAINT FK_Competitors_Students REFERENCES dbo.Students(Id),
        SuggestedStudentId int NULL,
        LinkStatus nvarchar(20) NOT NULL
    );
END;
