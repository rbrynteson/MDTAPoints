using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using MDTAPoints;

// Top-level statements
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Use NonCommercial for non-profit use

// Ensure CSV files exist (create empty ones if they don't)
EnsureCSVFilesExist();

// Main menu loop
while (true)
{
    Console.Clear();
    Console.WriteLine("MDTA Points Calculator");
    Console.WriteLine("======================");
    Console.WriteLine("Choose an option:");
    Console.WriteLine("1. Process existing event data (analyze CSV files)");
    Console.WriteLine("2. Add tournament data from Tabroom.com (web scraping)");
    Console.WriteLine("3. Generate markdown report for all events");
    Console.WriteLine("4. Generate detailed markdown report with tournament breakdowns");
    Console.WriteLine("5. Generate CSV file for an event");
    Console.WriteLine("6. Exit");
    Console.Write("Enter choice (1, 2, 3, 4, 5, or 6): ");

    string choice = Console.ReadLine()?.Trim() ?? "1";

    if (choice == "6")
    {
        Console.WriteLine("Goodbye!");
        break;
    }

if (choice == "3")
{
    // Generate comprehensive markdown report
    Console.WriteLine("\nGenerating Markdown Report for All Events");
    Console.WriteLine("=========================================");
    
    try
    {
        string report = GenerateMarkdownReport();
        
        // Save to file
        string reportFileName = $"MDTA_Report_{DateTime.Now:yyyy-MM-dd_HH-mm}.md";
        File.WriteAllText(reportFileName, report);
        
        Console.WriteLine($"Report generated: {reportFileName}");
        Console.WriteLine("\nReport content (copy/paste ready for Notion):");
        Console.WriteLine(new string('=', 60));
        Console.WriteLine(report);
        Console.WriteLine(new string('=', 60));
        Console.WriteLine($"\nReport also saved to: {reportFileName}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error generating report: {ex.Message}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
}

if (choice == "4")
{
    // Generate detailed markdown report with tournament breakdowns
    Console.WriteLine("\nGenerating Detailed Markdown Report for All Events");
    Console.WriteLine("==================================================");
    
    try
    {
        string report = GenerateDetailedMarkdownReport();
        
        // Save to file
        string reportFileName = $"MDTA_Detailed_Report_{DateTime.Now:yyyy-MM-dd_HH-mm}.md";
        File.WriteAllText(reportFileName, report);
        
        Console.WriteLine($"Detailed report generated: {reportFileName}");
        Console.WriteLine("\nDetailed report content (copy/paste ready for Notion):");
        Console.WriteLine(new string('=', 60));
        Console.WriteLine(report);
        Console.WriteLine(new string('=', 60));
        Console.WriteLine($"\nDetailed report also saved to: {reportFileName}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error generating detailed report: {ex.Message}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
}

if (choice == "5")
{
    // Generate CSV file for an event
    Console.WriteLine("\nGenerate CSV File for Event");
    Console.WriteLine("===========================");
    Console.WriteLine("Available event files:");

    var csvEventFiles = new Dictionary<string, string>
    {
        ["1"] = "ClassicDebate.csv",
        ["2"] = "CODebate.csv", 
        ["3"] = "LDDebate.csv",
        ["4"] = "PFDebate.csv",
        ["5"] = "CXDebate.csv"
    };

    foreach (var kvp in csvEventFiles)
    {
        var fileExists = File.Exists(kvp.Value);
        var status = fileExists ? "✓" : "✗";
        Console.WriteLine($"{kvp.Key}. {kvp.Value} {status}");
    }

    Console.Write("Select event file to export (1-5): ");
    string csvFileChoice = Console.ReadLine()?.Trim() ?? "1";

    if (!csvEventFiles.ContainsKey(csvFileChoice))
    {
        Console.WriteLine("Invalid choice.");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }

    string csvFilePath = csvEventFiles[csvFileChoice];

    if (!File.Exists(csvFilePath))
    {
        Console.WriteLine($"File {csvFilePath} not found. Please add tournament data first using option 2.");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }

    try
    {
        string outputCsvFileName = GenerateEventCSV(csvFilePath);
        Console.WriteLine($"\nCSV file generated successfully: {outputCsvFileName}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error generating CSV: {ex.Message}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
}

if (choice == "2")
{
    // Web scraping option - Add tournament data
    Console.WriteLine("\nAdd Tournament Data from Tabroom.com");
    Console.WriteLine("====================================");
    
    // Show event file selection menu
    Console.WriteLine("Select event type to add tournament data to:");
    Console.WriteLine("1. ClassicDebate.csv - Classic debate tournaments");
    Console.WriteLine("2. CODebate.csv - Congressional debate tournaments");
    Console.WriteLine("3. LDDebate.csv - Lincoln-Douglas debate tournaments");
    Console.WriteLine("4. PFDebate.csv - Public Forum debate tournaments");
    Console.WriteLine("5. CXDebate.csv - Cross-Examination/Policy debate tournaments");
    Console.Write("Select event type (1-5): ");
    
    string eventChoice = Console.ReadLine()?.Trim() ?? "1";
    
    var eventMapping = new Dictionary<string, (string fileName, string eventType)>
    {
        ["1"] = ("ClassicDebate.csv", "Classic"),
        ["2"] = ("CODebate.csv", "Congress"),
        ["3"] = ("LDDebate.csv", "LD"),
        ["4"] = ("PFDebate.csv", "PF"),
        ["5"] = ("CXDebate.csv", "CX")
    };
    
    if (!eventMapping.ContainsKey(eventChoice))
    {
        Console.WriteLine("Invalid choice.");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
    
    var (fileName, eventType) = eventMapping[eventChoice];
    Console.WriteLine($"Selected: {fileName}");
    
    Console.Write("\nEnter Tabroom.com results URL: ");
    string url = Console.ReadLine()?.Trim() ?? "";
    
    Console.Write("Enter tournament name: ");
    string tournamentName = Console.ReadLine()?.Trim() ?? "Tournament";

    if (string.IsNullOrEmpty(url))
    {
        Console.WriteLine("URL is required.");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }

    try
    {
        using var scraper = new TabroomScraper();
        var entries = await scraper.ScrapeTabroomResults(url, tournamentName, eventType);
        
        if (entries.Count == 0)
        {
            Console.WriteLine("No data found or error occurred during scraping.");
            Console.WriteLine("Press any key to return to main menu.");
            Console.ReadKey();
            continue;
        }

        // Append to appropriate event CSV file
        await scraper.AppendToEventCSV(entries, eventType);
        
        Console.WriteLine($"Tournament data added successfully!");
        Console.WriteLine("You can now run option 1 to analyze the updated data.");
        Console.WriteLine("\nPress any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error during web scraping: {ex.Message}");
        Console.WriteLine("Press any key to return to main menu.");
        Console.ReadKey();
        continue;
    }
}

// Option 1: Process existing event data
Console.WriteLine("\nProcess Event Data");
Console.WriteLine("==================");
Console.WriteLine("Available event files:");

var eventFiles = new Dictionary<string, string>
{
    ["1"] = "ClassicDebate.csv",
    ["2"] = "CODebate.csv", 
    ["3"] = "LDDebate.csv",
    ["4"] = "PFDebate.csv",
    ["5"] = "CXDebate.csv"
};

foreach (var kvp in eventFiles)
{
    var fileExists = File.Exists(kvp.Value);
    var status = fileExists ? "✓" : "✗";
    Console.WriteLine($"{kvp.Key}. {kvp.Value} {status}");
}

Console.Write("Select event file to process (1-5): ");
string fileChoice = Console.ReadLine()?.Trim() ?? "1";

if (!eventFiles.ContainsKey(fileChoice))
{
    Console.WriteLine("Invalid choice.");
    Console.WriteLine("Press any key to return to main menu.");
    Console.ReadKey();
    continue;
}

string filePath = eventFiles[fileChoice];

if (!File.Exists(filePath))
{
    Console.WriteLine($"File {filePath} not found. Please add tournament data first using option 2.");
    Console.WriteLine("Press any key to return to main menu.");
    Console.ReadKey();
    continue;
}

// Define dictionaries to hold individual tournament scores
var tournamentScores = new List<EntryTournamentScore>();

// Load CSV file with new format
using (var reader = new StreamReader(filePath))
{
    string? headerLine = reader.ReadLine(); // Skip header
    string? line;
    
    while ((line = reader.ReadLine()) != null)
    {
        var parts = line.Split(',');
        if (parts.Length >= 6)
        {
            string tournament = parts[0].Trim();
            string year = parts[1].Trim();
            string placeText = parts[2].Trim();
            string entry = parts[3].Trim();
            string school = parts[4].Trim();
            string elimPointsText = parts[5].Trim();

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
                // Add this entry's score for the tournament - start with just participation point
            tournamentScores.Add(new EntryTournamentScore
            {
                Entry = $"{school}:{name}",
                School = school,
                Tournament = tournament,
                Year = year,
                    Points = 1, // Start with just participation point - bonuses added later
                ElimPoints = elimPoints,
                Place = place
            });
            }
        }
    }
}

// ======= BONUS CALCULATION ========

// Determine if this is a Congress event
bool isCongress = filePath.Contains("CODebate", StringComparison.OrdinalIgnoreCase);

// Group scores by tournament
var groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();

// Track tournaments that don't meet requirements
var excludedTournaments = new List<string>();
var smallTournaments = new List<string>();

// First pass: Check which tournaments meet the 4-school minimum
var validTournaments = new HashSet<string>();
foreach (var group in groupedScores)
{
    var tournamentEntries = group.ToList();
    var uniqueSchools = tournamentEntries.Select(e => e.School).Distinct().Count();
    
    if (uniqueSchools >= 4)
    {
        validTournaments.Add(group.Key);
    }
    else
    {
        excludedTournaments.Add($"{group.Key} (only {uniqueSchools} school{(uniqueSchools == 1 ? "" : "s")})");
    }
}

// Remove entries from tournaments that don't meet the 4-school requirement
tournamentScores = tournamentScores.Where(ts => validTournaments.Contains(ts.Tournament)).ToList();

// Re-group after filtering
groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();

// Process each tournament to calculate bonuses
foreach (var group in groupedScores)
{
    var tournamentEntries = group.ToList();
    int totalEntries = tournamentEntries.Count;

    // Track if this is a small tournament for notes
    if (totalEntries <= 11)
    {
        smallTournaments.Add($"{group.Key} ({totalEntries} entries)");
    }
    
    foreach (var entry in tournamentEntries)
    {
        // Calculate bonus based on event type and placement
        int bonusPoints = 0;

        
        if (isCongress)
        {
            // Congress-specific scoring rules
            if (entry.Place <= 12) // Made finals (assuming top 12 make finals)
            {
                bonusPoints += 1; // Finals bonus
                
                if (entry.Place == 1)
                    bonusPoints += 2; // 1st place gets additional 2 points (total +3)
                else if (entry.Place >= 2 && entry.Place <= 8)
                    bonusPoints += 1; // Top 8 gets additional 1 point (total +2)
                // Places 9-12 get only finals bonus (+1)
            }
            // Prelim-only participants (place > 100) get no bonus points
        }
        else
        {
            // Standard scoring for Classic, LD, PF, CX
            if (totalEntries <= 11)
            {
                // For tournaments with 11 or fewer entries: only half get bonus, no first place points
                int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
            }
            else if (totalEntries == 12)
            {
                // For tournaments with exactly 12 entries, award 1 point to top half
                int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
            }
            else
            {
                // Standard bonus logic for larger tournaments (>12 entries)
                if (entry.Place == 1)
                    bonusPoints = 2;
                else if (entry.Place >= 2 && entry.Place <= 8)
                    bonusPoints = 1;
            }
        }

        // Update the entry's points with bonus points and elimination points
        entry.Points += bonusPoints;
        entry.Points += entry.ElimPoints;
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

// Show scoring explanation
Console.WriteLine("\n" + new string('=', 80));
Console.WriteLine("SCORING BREAKDOWN");
Console.WriteLine(new string('=', 80));
Console.WriteLine("Points awarded per tournament:");
Console.WriteLine("• Participation: 1 point");

if (isCongress)
{
    Console.WriteLine("• Congress-specific rules:");
    Console.WriteLine("  - Finals qualification: +1 point (top 12)");
    Console.WriteLine("  - Top 8 finish: +1 additional point");
    Console.WriteLine("  - 1st place: +2 additional points");
    Console.WriteLine("  - Maximum: 4 points (1 + 1 + 1 + 1)");
}
else
{
    Console.WriteLine("• Small tournaments (≤11 entries): +1 point for top half (no first place bonus)");
    Console.WriteLine("• Medium tournaments (12 entries): +1 point for top half");
    Console.WriteLine("• Large tournaments (>12 entries): +2 points for 1st, +1 point for 2nd-8th");
    Console.WriteLine("• Maximum: 3 points (1 + 2)");
}
Console.WriteLine();
Console.WriteLine("Tournament Requirements:");
Console.WriteLine("• Minimum 4 schools must attend for a tournament to count");
Console.WriteLine();

// Generate student scores output with better formatting
Console.WriteLine(new string('=', 80));
Console.WriteLine("STUDENT SCORES");
Console.WriteLine(new string('=', 80));

// Create formatted table
var tournaments = tournamentScores.Select(s => s.Tournament).Distinct().ToList();
Console.WriteLine($"{"Rank",-4} {"Student",-25} {"School",-20} {"Tournaments",-12} {"Total",-6}");
Console.WriteLine(new string('-', 80));

int rank = 1;
foreach (var entry in entryScores)
{
    var studentName = entry.Entry.Split(':').LastOrDefault() ?? entry.Entry;
    Console.WriteLine($"{rank,-4} {studentName,-25} {entry.School,-20} {entry.TotalTournaments,-12} {entry.TotalPoints,-6}");
    
    // Show tournament breakdown
    foreach (var tournament in tournaments)
    {
        var points = entry.TournamentPoints.GetValueOrDefault(tournament, 0);
        if (points > 0)
        {
            Console.WriteLine($"     └─ {tournament}: {points} points");
        }
    }
    Console.WriteLine();
    rank++;
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

// Generate school scores output with better formatting
Console.WriteLine("\n" + new string('=', 80));
Console.WriteLine("SCHOOL SCORES (Top 2 students per tournament)");
Console.WriteLine(new string('=', 80));

Console.WriteLine($"{"Rank",-4} {"School",-30} {"Total Points",-12}");
Console.WriteLine(new string('-', 80));

int schoolRank = 1;
foreach (var school in schoolTotals)
{
    Console.WriteLine($"{schoolRank,-4} {school.School,-30} {school.TotalPoints,-12}");
    
    // Show tournament breakdown
    foreach (var tournament in tournaments)
    {
        var points = school.TournamentScores.GetValueOrDefault(tournament, 0);
        if (points > 0)
        {
            Console.WriteLine($"     └─ {tournament}: {points} points (top 2 students)");
        }
    }
    Console.WriteLine();
    schoolRank++;
}

// Display notes about excluded and small tournaments
if (excludedTournaments.Any() || smallTournaments.Any())
{
    Console.WriteLine("\n" + new string('=', 80));
    Console.WriteLine("NOTES");
    Console.WriteLine(new string('=', 80));
    
    if (excludedTournaments.Any())
    {
        Console.WriteLine("\nTournaments excluded (fewer than 4 schools):");
        foreach (var tournament in excludedTournaments)
        {
            Console.WriteLine($"  • {tournament}");
        }
    }
    
    if (smallTournaments.Any())
    {
        Console.WriteLine("\nSmall tournaments (≤11 entries - modified scoring):");
        foreach (var tournament in smallTournaments)
        {
            Console.WriteLine($"  • {tournament}");
        }
        Console.WriteLine("  Note: Only top half receive bonus points; no first place bonus awarded");
    }
}

Console.WriteLine("\nProcessing complete. Press any key to return to main menu.");
Console.ReadKey();
}

// Method to ensure CSV files exist with headers
static void EnsureCSVFilesExist()
{
    var csvFiles = new[]
    {
        "ClassicDebate.csv",
        "CODebate.csv", 
        "LDDebate.csv",
        "PFDebate.csv",
        "CXDebate.csv"
    };

    string header = "Tournament Name,Year,Place,Entry,School,ElimPoints";

    foreach (var fileName in csvFiles)
    {
        if (!File.Exists(fileName))
        {
            Console.WriteLine($"Creating empty {fileName}...");
            File.WriteAllText(fileName, header + Environment.NewLine);
        }
    }
}

// Method to generate comprehensive markdown report for all events
static string GenerateMarkdownReport()
{
    var report = new System.Text.StringBuilder();
    
    // Report header
    report.AppendLine("# MDTA Tournament Results Summary");
    report.AppendLine($"*Generated on {DateTime.Now:MMMM dd, yyyy} at {DateTime.Now:h:mm tt}*");
    report.AppendLine();

    var eventFiles = new Dictionary<string, string>
    {
        ["Classic Debate"] = "ClassicDebate.csv",
        ["Congressional Debate"] = "CODebate.csv",
        ["Lincoln-Douglas Debate"] = "LDDebate.csv",
        ["Public Forum Debate"] = "PFDebate.csv",
        ["Cross-Examination Debate"] = "CXDebate.csv"
    };

    foreach (var eventPair in eventFiles)
    {
        string eventName = eventPair.Key;
        string fileName = eventPair.Value;
        
        if (File.Exists(fileName))
        {
            var eventData = ProcessEventForReport(fileName, eventName);
            if (eventData.HasData)
            {
                report.AppendLine($"## {eventName}");
                report.AppendLine();
                
                // Top Schools
                report.AppendLine("### Top Schools");
                if (eventData.TopSchools.Any())
                {
                    for (int i = 0; i < eventData.TopSchools.Count; i++)
                    {
                        var school = eventData.TopSchools[i];
                        string medal = i == 0 ? "1st" : i == 1 ? "2nd" : "3rd";
                        report.AppendLine($"**{medal} Place: {school.School}** - {school.Points} points ({school.Tournaments} tournaments)");
                    }
                }
                else
                {
                    report.AppendLine("*No tournament data available*");
                }
                report.AppendLine();
                
                // Top Students
                report.AppendLine("### Top Students");
                if (eventData.TopStudents.Any())
                {
                    report.AppendLine("| Rank | Student | School | Points | Tournaments |");
                    report.AppendLine("|------|---------|--------|--------|-------------|");
                    
                    // Calculate proper rankings with ties
                    int currentRank = 1;
                    for (int i = 0; i < eventData.TopStudents.Count; i++)
                    {
                        var student = eventData.TopStudents[i];
                        string studentName = student.Entry.Split(':').LastOrDefault() ?? student.Entry;
                        
                        // Check if this student is tied with others
                        var studentsWithSamePoints = eventData.TopStudents.Where(s => s.Points == student.Points).ToList();
                        string rankDisplay;
                        
                        if (studentsWithSamePoints.Count > 1)
                        {
                            // Multiple students with same points - show as tie
                            rankDisplay = $"T-{currentRank}";
                        }
                        else
                        {
                            // Single student at this point level
                            rankDisplay = currentRank.ToString();
                        }
                        
                        report.AppendLine($"| {rankDisplay} | {studentName} | {student.School} | {student.Points} | {student.Tournaments} |");
                        
                        // Update rank for next iteration
                        if (i + 1 < eventData.TopStudents.Count && eventData.TopStudents[i + 1].Points != student.Points)
                        {
                            // Next student has different points, so rank jumps by number of tied students
                            currentRank = i + 2;
                        }
                    }
                }
                else
                {
                    report.AppendLine("*No tournament data available*");
                }
                report.AppendLine();
                
                // Add notes section if there are any excluded or small tournaments
                if (eventData.ExcludedTournaments.Any() || eventData.SmallTournaments.Any())
                {
                    report.AppendLine("### Notes");
                    
                    if (eventData.ExcludedTournaments.Any())
                    {
                        report.AppendLine("**Tournaments excluded (fewer than 4 schools):**");
                        foreach (var tournament in eventData.ExcludedTournaments)
                        {
                            report.AppendLine($"- {tournament}");
                        }
                        report.AppendLine();
                    }
                    
                    if (eventData.SmallTournaments.Any())
                    {
                        report.AppendLine("**Small tournaments (≤11 entries - modified scoring):**");
                        foreach (var tournament in eventData.SmallTournaments)
                        {
                            report.AppendLine($"- {tournament}");
                        }
                        report.AppendLine("*Note: Only top half receive bonus points; no first place bonus awarded*");
                        report.AppendLine();
                    }
                }
            }
        }
    }
    
    // Footer
    report.AppendLine("---");
    report.AppendLine("*Report generated by MDTA Points Calculator*");
    
    return report.ToString();
}

// Method to generate detailed markdown report with tournament breakdowns
static string GenerateDetailedMarkdownReport()
{
    var report = new System.Text.StringBuilder();
    
    // Report header
    report.AppendLine("# MDTA Tournament Results Summary (Detailed)");
    report.AppendLine($"*Generated on {DateTime.Now:MMMM dd, yyyy} at {DateTime.Now:h:mm tt}*");
    report.AppendLine();

    var eventFiles = new Dictionary<string, string>
    {
        ["Classic Debate"] = "ClassicDebate.csv",
        ["Congressional Debate"] = "CODebate.csv",
        ["Lincoln-Douglas Debate"] = "LDDebate.csv",
        ["Public Forum Debate"] = "PFDebate.csv",
        ["Cross-Examination Debate"] = "CXDebate.csv"
    };

    foreach (var eventPair in eventFiles)
    {
        string eventName = eventPair.Key;
        string fileName = eventPair.Value;
        
        if (File.Exists(fileName))
        {
            var eventData = ProcessEventForDetailedReport(fileName, eventName);
            if (eventData.HasData)
            {
                report.AppendLine($"## {eventName}");
                report.AppendLine();
                
                // Top Schools with tournament breakdown
                report.AppendLine("### Top Schools");
                if (eventData.TopSchools.Any())
                {
                    for (int i = 0; i < eventData.TopSchools.Count; i++)
                    {
                        var school = eventData.TopSchools[i];
                        string medal = i == 0 ? "1st" : i == 1 ? "2nd" : "3rd";
                        report.AppendLine($"**{medal} Place: {school.School}** - {school.Points} points ({school.Tournaments} tournaments)");
                        
                        // Add tournament breakdown for schools
                        foreach (var tournamentScore in school.TournamentBreakdown)
                        {
                            if (tournamentScore.Value > 0)
                            {
                                report.AppendLine($"  - {tournamentScore.Key}: {tournamentScore.Value} points (top 2 students)");
                            }
                        }
                        report.AppendLine();
                    }
                }
                else
                {
                    report.AppendLine("*No tournament data available*");
                }
                report.AppendLine();
                
                // Top Students with tournament breakdown
                report.AppendLine("### Top Students");
                if (eventData.TopStudents.Any())
                {
                    report.AppendLine("| Rank | Student | School | Points | Tournaments |");
                    report.AppendLine("|------|---------|--------|--------|-------------|");
                    
                    // Calculate proper rankings with ties
                    int currentRank = 1;
                    for (int i = 0; i < eventData.TopStudents.Count; i++)
                    {
                        var student = eventData.TopStudents[i];
                        string studentName = student.Entry.Split(':').LastOrDefault() ?? student.Entry;
                        
                        // Check if this student is tied with others
                        var studentsWithSamePoints = eventData.TopStudents.Where(s => s.Points == student.Points).ToList();
                        string rankDisplay;
                        
                        if (studentsWithSamePoints.Count > 1)
                        {
                            // Multiple students with same points - show as tie
                            rankDisplay = $"T-{currentRank}";
                        }
                        else
                        {
                            // Single student at this point level
                            rankDisplay = currentRank.ToString();
                        }
                        
                        report.AppendLine($"| {rankDisplay} | {studentName} | {student.School} | {student.Points} | {student.Tournaments} |");
                        
                        // Add tournament breakdown for students (spanning Student and School columns)
                        foreach (var tournamentScore in student.TournamentBreakdown)
                        {
                            if (tournamentScore.Value > 0)
                            {
                                report.AppendLine($"| | └─ {tournamentScore.Key}: | {tournamentScore.Value} points | | |");
                            }
                        }
                        
                        // Update rank for next iteration
                        if (i + 1 < eventData.TopStudents.Count && eventData.TopStudents[i + 1].Points != student.Points)
                        {
                            // Next student has different points, so rank jumps by number of tied students
                            currentRank = i + 2;
                        }
                    }
                }
                else
                {
                    report.AppendLine("*No tournament data available*");
                }
                report.AppendLine();
                
                // Add notes section if there are any excluded or small tournaments
                if (eventData.ExcludedTournaments.Any() || eventData.SmallTournaments.Any())
                {
                    report.AppendLine("### Notes");
                    
                    if (eventData.ExcludedTournaments.Any())
                    {
                        report.AppendLine("**Tournaments excluded (fewer than 4 schools):**");
                        foreach (var tournament in eventData.ExcludedTournaments)
                        {
                            report.AppendLine($"- {tournament}");
                        }
                        report.AppendLine();
                    }
                    
                    if (eventData.SmallTournaments.Any())
                    {
                        report.AppendLine("**Small tournaments (≤11 entries - modified scoring):**");
                        foreach (var tournament in eventData.SmallTournaments)
                        {
                            report.AppendLine($"- {tournament}");
                        }
                        report.AppendLine("*Note: Only top half receive bonus points; no first place bonus awarded*");
                        report.AppendLine();
                    }
                }
            }
        }
    }
    
    // Footer
    report.AppendLine("---");
    report.AppendLine("*Detailed report generated by MDTA Points Calculator*");
    
    return report.ToString();
}

// Helper method to process event data for detailed reporting
static EventReportData ProcessEventForDetailedReport(string filePath, string eventName)
{
    var tournamentScores = new List<EntryTournamentScore>();
    
    // Load CSV file (same as regular report)
    try
    {
        using (var reader = new StreamReader(filePath))
        {
            string? headerLine = reader.ReadLine(); // Skip header
            string? line;
            
            while ((line = reader.ReadLine()) != null)
            {
                var parts = line.Split(',');
                if (parts.Length >= 6)
                {
                    string tournament = parts[0].Trim();
                    string year = parts[1].Trim();
                    string placeText = parts[2].Trim();
                    string entry = parts[3].Trim();
                    string school = parts[4].Trim();
                    string elimPointsText = parts[5].Trim();

                    // Skip invalid rows
                    if (string.IsNullOrEmpty(tournament) || string.IsNullOrEmpty(year) || string.IsNullOrEmpty(entry) || string.IsNullOrEmpty(school))
                        continue;

                    // Parse place and elimination points
                    int place = int.TryParse(placeText, out int parsedPlace) ? parsedPlace : 0;
                    int elimPoints = int.TryParse(elimPointsText, out int parsedElimPoints) ? parsedElimPoints : 0;

                    // Split entry into individual names
                    var names = entry.Split('&', StringSplitOptions.TrimEntries);

                    foreach (var name in names)
                    {
                        tournamentScores.Add(new EntryTournamentScore
                        {
                            Entry = $"{school}:{name}",
                            School = school,
                            Tournament = tournament,
                            Year = year,
                            Points = 1, // Start with participation point
                            ElimPoints = elimPoints,
                            Place = place
                        });
                    }
                }
            }
        }
    }
    catch
    {
        return new EventReportData { HasData = false };
    }

    if (!tournamentScores.Any())
    {
        return new EventReportData { HasData = false };
    }

    // Apply scoring logic (same as main program)
    bool isCongress = filePath.Contains("CODebate", StringComparison.OrdinalIgnoreCase);
    
    var groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    // Track tournaments that don't meet requirements
    var excludedTournaments = new List<string>();
    var smallTournaments = new List<string>();
    
    // First pass: Check which tournaments meet the 4-school minimum
    var validTournaments = new HashSet<string>();
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        var uniqueSchools = tournamentEntries.Select(e => e.School).Distinct().Count();
        
        if (uniqueSchools >= 4)
        {
            validTournaments.Add(group.Key);
        }
        else
        {
            excludedTournaments.Add($"{group.Key} (only {uniqueSchools} school{(uniqueSchools == 1 ? "" : "s")})");
        }
    }
    
    // Remove entries from tournaments that don't meet the 4-school requirement
    tournamentScores = tournamentScores.Where(ts => validTournaments.Contains(ts.Tournament)).ToList();
    
    // Re-group after filtering
    groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        int totalEntries = tournamentEntries.Count;
        
        // Track if this is a small tournament for notes
        if (totalEntries <= 11)
        {
            smallTournaments.Add($"{group.Key} ({totalEntries} entries)");
        }

        foreach (var entry in tournamentEntries)
        {
            int bonusPoints = 0;
            
            if (isCongress)
            {
                if (entry.Place <= 12)
                {
                    bonusPoints += 1; // Finals bonus
                    
                    if (entry.Place == 1)
                        bonusPoints += 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints += 1;
                }
            }
            else
            {
                if (totalEntries <= 11)
                {
                    // For tournaments with 11 or fewer entries: only half get bonus, no first place points
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else if (totalEntries == 12)
                {
                    // For tournaments with exactly 12 entries, award 1 point to top half
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else
                {
                    // Standard bonus logic for larger tournaments (>12 entries)
                    if (entry.Place == 1)
                        bonusPoints = 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints = 1;
                }
            }

            entry.Points += bonusPoints;
            entry.Points += entry.ElimPoints;
        }
    }

    // Calculate student scores with tournament breakdown
    var studentScores = tournamentScores
        .GroupBy(score => score.Entry)
        .Select(group => new StudentReportData
        {
            Entry = group.Key,
            School = group.First().School,
            Points = group.Sum(score => score.Points),
            Tournaments = group.Select(score => score.Tournament).Distinct().Count(),
            TournamentBreakdown = group.GroupBy(score => score.Tournament)
                                      .ToDictionary(tGroup => tGroup.Key, tGroup => tGroup.Sum(s => s.Points))
        })
        .OrderByDescending(entry => entry.Points)
        .ThenBy(entry => entry.Entry)
        .ToList();

    // Handle ties for top 12 students (same logic as before)
    var topStudents = new List<StudentReportData>();
    if (studentScores.Any())
    {
        topStudents.Add(studentScores[0]); // Always include 1st place
        
        for (int i = 1; i < studentScores.Count && topStudents.Count < 12; i++)
        {
            topStudents.Add(studentScores[i]);
        }
        
        // Add tied students beyond 12th place
        if (topStudents.Count == 12 && topStudents.Count < studentScores.Count)
        {
            int lastPoints = topStudents.Last().Points;
            for (int i = 12; i < studentScores.Count; i++)
            {
                if (studentScores[i].Points == lastPoints)
                {
                    topStudents.Add(studentScores[i]);
                }
                else
                {
                    break;
                }
            }
        }
    }

    // Calculate school scores with tournament breakdown
    var schoolScores = new Dictionary<string, Dictionary<string, int>>();
    foreach (var school in tournamentScores.GroupBy(score => score.School))
    {
        var schoolName = school.Key;
        schoolScores[schoolName] = new Dictionary<string, int>();

        foreach (var tournament in tournamentScores.Select(s => s.Tournament).Distinct())
        {
            var topTwoScores = school
                .Where(score => score.Tournament == tournament)
                .OrderByDescending(score => score.Points)
                .Take(2)
                .Sum(score => score.Points);

            schoolScores[schoolName][tournament] = topTwoScores;
        }
    }

    var schoolTotals = schoolScores
        .Select(school => new SchoolReportData
        {
            School = school.Key,
            Points = school.Value.Values.Sum(),
            Tournaments = school.Value.Count(kvp => kvp.Value > 0),
            TournamentBreakdown = school.Value
        })
        .OrderByDescending(s => s.Points)
        .ToList();

    // Handle ties for top 3 schools (same logic as before)
    var topSchools = new List<SchoolReportData>();
    if (schoolTotals.Any())
    {
        topSchools.Add(schoolTotals[0]); // Always include 1st place
        
        for (int i = 1; i < schoolTotals.Count && topSchools.Count < 3; i++)
        {
            topSchools.Add(schoolTotals[i]);
        }
        
        // Add tied schools beyond 3rd place
        if (topSchools.Count == 3 && topSchools.Count < schoolTotals.Count)
        {
            int lastPoints = topSchools.Last().Points;
            for (int i = 3; i < schoolTotals.Count; i++)
            {
                if (schoolTotals[i].Points == lastPoints)
                {
                    topSchools.Add(schoolTotals[i]);
                }
                else
                {
                    break;
                }
            }
        }
    }

    return new EventReportData
    {
        HasData = true,
        TopStudents = topStudents,
        TopSchools = topSchools,
        ExcludedTournaments = excludedTournaments,
        SmallTournaments = smallTournaments
    };
}

// Method to generate CSV file for an event
static string GenerateEventCSV(string filePath)
{
    var tournamentScores = new List<EntryTournamentScore>();
    
    // Load CSV file
    using (var reader = new StreamReader(filePath))
    {
        string? headerLine = reader.ReadLine(); // Skip header
        string? line;
        
        while ((line = reader.ReadLine()) != null)
        {
            var parts = line.Split(',');
            if (parts.Length >= 6)
            {
                string tournament = parts[0].Trim();
                string year = parts[1].Trim();
                string placeText = parts[2].Trim();
                string entry = parts[3].Trim();
                string school = parts[4].Trim();
                string elimPointsText = parts[5].Trim();

                // Skip invalid rows
                if (string.IsNullOrEmpty(tournament) || string.IsNullOrEmpty(year) || string.IsNullOrEmpty(entry) || string.IsNullOrEmpty(school))
                    continue;

                // Parse place and elimination points
                int place = int.TryParse(placeText, out int parsedPlace) ? parsedPlace : 0;
                int elimPoints = int.TryParse(elimPointsText, out int parsedElimPoints) ? parsedElimPoints : 0;

                // Split entry into individual names
                var names = entry.Split('&', StringSplitOptions.TrimEntries);

                foreach (var name in names)
                {
                    tournamentScores.Add(new EntryTournamentScore
                    {
                        Entry = $"{school}:{name}",
                        School = school,
                        Tournament = tournament,
                        Year = year,
                        Points = 1, // Start with participation point
                        ElimPoints = elimPoints,
                        Place = place
                    });
                }
            }
        }
    }

    // Apply scoring logic (same as main program)
    bool isCongress = filePath.Contains("CODebate", StringComparison.OrdinalIgnoreCase);
    
    var groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    // First pass: Check which tournaments meet the 4-school minimum
    var validTournaments = new HashSet<string>();
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        var uniqueSchools = tournamentEntries.Select(e => e.School).Distinct().Count();
        
        if (uniqueSchools >= 4)
        {
            validTournaments.Add(group.Key);
        }
    }
    
    // Remove entries from tournaments that don't meet the 4-school requirement
    tournamentScores = tournamentScores.Where(ts => validTournaments.Contains(ts.Tournament)).ToList();
    
    // Re-group after filtering
    groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    // Process each tournament to calculate bonuses
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        int totalEntries = tournamentEntries.Count;
        
        foreach (var entry in tournamentEntries)
        {
            int bonusPoints = 0;
            
            if (isCongress)
            {
                if (entry.Place <= 12)
                {
                    bonusPoints += 1; // Finals bonus
                    
                    if (entry.Place == 1)
                        bonusPoints += 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints += 1;
                }
            }
            else
            {
                if (totalEntries <= 11)
                {
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else if (totalEntries == 12)
                {
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else
                {
                    if (entry.Place == 1)
                        bonusPoints = 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints = 1;
                }
            }

            entry.Points += bonusPoints;
            entry.Points += entry.ElimPoints;
        }
    }

    // Calculate student scores
    var studentScores = tournamentScores
        .GroupBy(score => score.Entry)
        .Select(group => new
        {
            Entry = group.Key,
            School = group.First().School,
            TotalTournaments = group.Select(score => score.Tournament).Distinct().Count(),
            TotalPoints = group.Sum(score => score.Points)
        })
        .OrderByDescending(entry => entry.TotalPoints)
        .ThenBy(entry => entry.Entry)
        .ToList();

    // Generate output CSV file name
    string eventName = Path.GetFileNameWithoutExtension(filePath);
    string outputFileName = $"{eventName}_Results_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv";

    // Write CSV file
    using (var writer = new StreamWriter(outputFileName))
    {
        // Write header
        writer.WriteLine("Rank,Student Name,School,Total Points,Number of Tournaments");
        
        // Write data
        int rank = 1;
        foreach (var student in studentScores)
        {
            string studentName = student.Entry.Split(':').LastOrDefault() ?? student.Entry;
            writer.WriteLine($"{rank},{studentName},{student.School},{student.TotalPoints},{student.TotalTournaments}");
            rank++;
        }
    }

    return outputFileName;
}

// Helper method to process event data for reporting
static EventReportData ProcessEventForReport(string filePath, string eventName)
{
    var tournamentScores = new List<EntryTournamentScore>();
    
    // Load CSV file
    try
    {
        using (var reader = new StreamReader(filePath))
        {
            string? headerLine = reader.ReadLine(); // Skip header
            string? line;
            
            while ((line = reader.ReadLine()) != null)
            {
                var parts = line.Split(',');
                if (parts.Length >= 6)
                {
                    string tournament = parts[0].Trim();
                    string year = parts[1].Trim();
                    string placeText = parts[2].Trim();
                    string entry = parts[3].Trim();
                    string school = parts[4].Trim();
                    string elimPointsText = parts[5].Trim();

                    // Skip invalid rows
                    if (string.IsNullOrEmpty(tournament) || string.IsNullOrEmpty(year) || string.IsNullOrEmpty(entry) || string.IsNullOrEmpty(school))
                        continue;

                    // Parse place and elimination points
                    int place = int.TryParse(placeText, out int parsedPlace) ? parsedPlace : 0;
                    int elimPoints = int.TryParse(elimPointsText, out int parsedElimPoints) ? parsedElimPoints : 0;

                    // Split entry into individual names
                    var names = entry.Split('&', StringSplitOptions.TrimEntries);

                    foreach (var name in names)
                    {
                        tournamentScores.Add(new EntryTournamentScore
                        {
                            Entry = $"{school}:{name}",
                            School = school,
                            Tournament = tournament,
                            Year = year,
                            Points = 1, // Start with participation point
                            ElimPoints = elimPoints,
                            Place = place
                        });
                    }
                }
            }
        }
    }
    catch
    {
        return new EventReportData { HasData = false };
    }

    if (!tournamentScores.Any())
    {
        return new EventReportData { HasData = false };
    }

    // Apply scoring logic (same as main program)
    bool isCongress = filePath.Contains("CODebate", StringComparison.OrdinalIgnoreCase);
    
    var groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    // Track tournaments that don't meet requirements
    var excludedTournaments = new List<string>();
    var smallTournaments = new List<string>();
    
    // First pass: Check which tournaments meet the 4-school minimum
    var validTournaments = new HashSet<string>();
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        var uniqueSchools = tournamentEntries.Select(e => e.School).Distinct().Count();
        
        if (uniqueSchools >= 4)
        {
            validTournaments.Add(group.Key);
        }
        else
        {
            excludedTournaments.Add($"{group.Key} (only {uniqueSchools} school{(uniqueSchools == 1 ? "" : "s")})");
        }
    }
    
    // Remove entries from tournaments that don't meet the 4-school requirement
    tournamentScores = tournamentScores.Where(ts => validTournaments.Contains(ts.Tournament)).ToList();
    
    // Re-group after filtering
    groupedScores = tournamentScores.GroupBy(ts => ts.Tournament).ToList();
    
    foreach (var group in groupedScores)
    {
        var tournamentEntries = group.ToList();
        int totalEntries = tournamentEntries.Count;
        
        // Track if this is a small tournament for notes
        if (totalEntries <= 11)
        {
            smallTournaments.Add($"{group.Key} ({totalEntries} entries)");
        }

        foreach (var entry in tournamentEntries)
        {
            int bonusPoints = 0;
            
            if (isCongress)
            {
                if (entry.Place <= 12)
                {
                    bonusPoints += 1; // Finals bonus
                    
                    if (entry.Place == 1)
                        bonusPoints += 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints += 1;
                }
            }
            else
            {
                if (totalEntries <= 11)
                {
                    // For tournaments with 11 or fewer entries: only half get bonus, no first place points
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else if (totalEntries == 12)
                {
                    // For tournaments with exactly 12 entries, award 1 point to top half
                    int topHalfCutoff = (int)Math.Ceiling(totalEntries / 2.0);
                    bonusPoints = entry.Place <= topHalfCutoff ? 1 : 0;
                }
                else
                {
                    // Standard bonus logic for larger tournaments (>12 entries)
                    if (entry.Place == 1)
                        bonusPoints = 2;
                    else if (entry.Place >= 2 && entry.Place <= 8)
                        bonusPoints = 1;
                }
            }

            entry.Points += bonusPoints;
            entry.Points += entry.ElimPoints;
        }
    }

    // Calculate student scores
    var studentScores = tournamentScores
        .GroupBy(score => score.Entry)
        .Select(group => new StudentReportData
        {
            Entry = group.Key,
            School = group.First().School,
            Points = group.Sum(score => score.Points),
            Tournaments = group.Select(score => score.Tournament).Distinct().Count()
        })
        .OrderByDescending(entry => entry.Points)
        .ThenBy(entry => entry.Entry)
        .ToList();

    // Handle ties for top 12 students
    var topStudents = new List<StudentReportData>();
    if (studentScores.Any())
    {
        topStudents.Add(studentScores[0]); // Always include 1st place
        
        for (int i = 1; i < studentScores.Count && topStudents.Count < 12; i++)
        {
            topStudents.Add(studentScores[i]);
        }
        
        // Add tied students beyond 12th place
        if (topStudents.Count == 12 && topStudents.Count < studentScores.Count)
        {
            int lastPoints = topStudents.Last().Points;
            for (int i = 12; i < studentScores.Count; i++)
            {
                if (studentScores[i].Points == lastPoints)
                {
                    topStudents.Add(studentScores[i]);
                }
                else
                {
                    break;
                }
            }
        }
    }

    // Calculate school scores
    var schoolScores = new Dictionary<string, Dictionary<string, int>>();
    foreach (var school in tournamentScores.GroupBy(score => score.School))
    {
        var schoolName = school.Key;
        schoolScores[schoolName] = new Dictionary<string, int>();

        foreach (var tournament in tournamentScores.Select(s => s.Tournament).Distinct())
        {
            var topTwoScores = school
                .Where(score => score.Tournament == tournament)
                .OrderByDescending(score => score.Points)
                .Take(2)
                .Sum(score => score.Points);

            schoolScores[schoolName][tournament] = topTwoScores;
        }
    }

    var schoolTotals = schoolScores
        .Select(school => new SchoolReportData
        {
            School = school.Key,
            Points = school.Value.Values.Sum(),
            Tournaments = school.Value.Count(kvp => kvp.Value > 0)
        })
        .OrderByDescending(s => s.Points)
        .ToList();

    // Handle ties for top 3 schools
    var topSchools = new List<SchoolReportData>();
    if (schoolTotals.Any())
    {
        topSchools.Add(schoolTotals[0]); // Always include 1st place
        
        for (int i = 1; i < schoolTotals.Count && topSchools.Count < 3; i++)
        {
            topSchools.Add(schoolTotals[i]);
        }
        
        // Add tied schools beyond 3rd place
        if (topSchools.Count == 3 && topSchools.Count < schoolTotals.Count)
        {
            int lastPoints = topSchools.Last().Points;
            for (int i = 3; i < schoolTotals.Count; i++)
            {
                if (schoolTotals[i].Points == lastPoints)
                {
                    topSchools.Add(schoolTotals[i]);
                }
                else
                {
                    break;
                }
            }
        }
    }

    return new EventReportData
    {
        HasData = true,
        TopStudents = topStudents,
        TopSchools = topSchools,
        ExcludedTournaments = excludedTournaments,
        SmallTournaments = smallTournaments
    };
}

// Helper classes for report data
class EventReportData
{
    public bool HasData { get; set; }
    public List<StudentReportData> TopStudents { get; set; } = new();
    public List<SchoolReportData> TopSchools { get; set; } = new();
    public List<string> ExcludedTournaments { get; set; } = new();
    public List<string> SmallTournaments { get; set; } = new();
}

class StudentReportData
{
    public string Entry { get; set; } = string.Empty;
    public string School { get; set; } = string.Empty;
    public int Points { get; set; }
    public int Tournaments { get; set; }
    public Dictionary<string, int> TournamentBreakdown { get; set; } = new();
}

class SchoolReportData
{
    public string School { get; set; } = string.Empty;
    public int Points { get; set; }
    public int Tournaments { get; set; }
    public Dictionary<string, int> TournamentBreakdown { get; set; } = new();
}

// Class to store individual tournament scores
class EntryTournamentScore
{
    public string Entry { get; set; } = string.Empty;
    public string School { get; set; } = string.Empty;
    public string Tournament { get; set; } = string.Empty;
    public string Year { get; set; } = string.Empty;
    public int Points { get; set; }
    public int ElimPoints { get; set; } // New field for ElimPoints
    public int Place { get; set; }
}
