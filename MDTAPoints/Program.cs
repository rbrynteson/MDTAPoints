using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

// Top-level statements
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Use NonCommercial for non-profit use

// Prompt the user for the Excel file path
Console.WriteLine("Enter the path to the Excel file:");
string filePath = Console.ReadLine();

if (!File.Exists(filePath))
{
    Console.WriteLine("File not found.");
    return;
}

// Define dictionaries to hold individual tournament scores
var tournamentScores = new List<EntryTournamentScore>();

// Load the Excel file
using (var package = new ExcelPackage(new FileInfo(filePath)))
{
    var worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension.Rows;

    for (int row = 2; row <= rowCount; row++) // Assuming row 1 is header
    {
        string tournament = worksheet.Cells[row, 1].Text.Trim();
        string year = worksheet.Cells[row, 3].Text.Trim();
        string placeText = worksheet.Cells[row, 4].Text.Trim();
        string entry = worksheet.Cells[row, 5].Text.Trim();
        string school = worksheet.Cells[row, 6].Text.Trim();
        string elimPointsText = worksheet.Cells[row, 9].Text.Trim(); // Read ElimPoints column

        // Skip invalid rows
        if (string.IsNullOrEmpty(tournament) || string.IsNullOrEmpty(year) || string.IsNullOrEmpty(entry) || string.IsNullOrEmpty(school))
            continue;

        // Parse place as an integer
        int place = int.TryParse(placeText, out int parsedPlace) ? parsedPlace : 0;

        // Parse ElimPoints as an integer (default to 0 if blank)
        int elimPoints = int.TryParse(elimPointsText, out int parsedElimPoints) ? parsedElimPoints : 0;

        // Split the Entry into individual names by '&'
        var names = entry.Split('&', StringSplitOptions.TrimEntries);

        foreach (var name in names)
        {
            // Add this entry's score for the tournament without calculating bonus
            tournamentScores.Add(new EntryTournamentScore
            {
                Entry = $"{school}:{name}",
                School = school,
                Tournament = tournament,
                Year = year,
                Points = 1 + elimPoints, // Participation always gives 1 point + ElimPoints
                ElimPoints = elimPoints,
                Place = place
            });
        }
    }
}

// ======= BONUS CALCULATION ========

// Group scores by tournament
var groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();

// Process each tournament to calculate bonuses
foreach (var group in groupedScores)
{
    var tournamentEntries = group.ToList();
    int totalEntries = tournamentEntries.Count;

    foreach (var entry in tournamentEntries)
    {
        // Calculate bonus based on the total number of entries
        int bonusPoints = 0;

        if (totalEntries <= 12)
        {
            // For tournaments with 12 or fewer entries, award 1 point to top half
            int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
            bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
        }
        else
        {
            // Standard bonus logic for larger tournaments
            if (entry.Place == 1)
                bonusPoints = 2;
            else if (entry.Place >= 2 && entry.Place <= 8)
                bonusPoints = 1;
        }

        // Update the entry's points
        entry.Points += bonusPoints;
    }
}

// ======== STUDENT SCORES OUTPUT ========

// Group by entry and calculate total points and total tournaments
var entryScores = tournamentScores
    .GroupBy(score => score.Entry)
    .Select(group => new
    {
        Entry = group.Key,
        School = group.First().School,
        TotalTournaments = group.Select(score => score.Tournament).Distinct().Count(), // Count unique tournaments
        TournamentPoints = group.GroupBy(score => score.Tournament)
                                .ToDictionary(score => score.Key, score => score.Sum(s => s.Points)),
        TotalPoints = group.Sum(score => score.Points)
    })
    .OrderByDescending(entry => entry.TotalPoints) // Sort by total points descending
    .ThenBy(entry => entry.Entry) // Secondary sort by name
    .ToList();

// Generate student scores output
Console.WriteLine("Student Scores:");
Console.WriteLine("Entry, School, Tournaments, " + string.Join(", ", tournamentScores.Select(s => s.Tournament).Distinct()) + ", Total Points");

foreach (var entry in entryScores)
{
    var tournamentColumns = string.Join(", ", tournamentScores.Select(s => s.Tournament).Distinct()
        .Select(tournament => entry.TournamentPoints.GetValueOrDefault(tournament, 0)));

    Console.WriteLine($"{entry.Entry}, {entry.School}, {entry.TotalTournaments}, {tournamentColumns}, {entry.TotalPoints}");
}

// ======== SCHOOL SCORES OUTPUT ========

// Group by school and calculate top 2 scores per tournament
var schoolScores = new Dictionary<string, Dictionary<string, int>>();

foreach (var school in tournamentScores.GroupBy(score => score.School))
{
    var schoolName = school.Key;
    schoolScores[schoolName] = new Dictionary<string, int>();

    foreach (var tournament in tournamentScores.Select(s => s.Tournament).Distinct())
    {
        // Get all scores for this school in this tournament
        var topTwoScores = school
            .Where(score => score.Tournament == tournament)
            .OrderByDescending(score => score.Points)
            .Take(2) // Take the top 2 scores
            .Sum(score => score.Points);

        // Add to school scores
        schoolScores[schoolName][tournament] = topTwoScores;
    }
}

// Generate school totals
var schoolTotals = schoolScores
    .Select(school => new
    {
        School = school.Key,
        TournamentScores = school.Value,
        TotalPoints = school.Value.Values.Sum() // Sum across all tournaments
    })
    .OrderByDescending(s => s.TotalPoints) // Sort by total points descending
    .ToList();

// Generate school scores output
Console.WriteLine("\nSchool Scores:");
Console.Write("School");
foreach (var tournament in tournamentScores.Select(s => s.Tournament).Distinct())
{
    Console.Write($", {tournament}");
}
Console.WriteLine(", Total Points");

foreach (var school in schoolTotals)
{
    Console.Write($"{school.School}");
    foreach (var tournament in tournamentScores.Select(s => s.Tournament).Distinct())
    {
        Console.Write($", {school.TournamentScores.GetValueOrDefault(tournament, 0)}");
    }
    Console.WriteLine($", {school.TotalPoints}");
}

Console.WriteLine("\nProcessing complete. Press any key to exit.");
Console.ReadKey();

// Class to store individual tournament scores
class EntryTournamentScore
{
    public string Entry { get; set; }
    public string School { get; set; }
    public string Tournament { get; set; }
    public string Year { get; set; }
    public int Points { get; set; }
    public int ElimPoints { get; set; } // New field for ElimPoints
    public int Place { get; set; }
}
