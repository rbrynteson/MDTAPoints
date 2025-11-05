using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace MDTAPoints
{
    public class TabroomScraper : IDisposable
    {
        private readonly HttpClient _httpClient;

        public TabroomScraper()
        {
            _httpClient = new HttpClient();
            // Set a user agent to avoid being blocked
            _httpClient.DefaultRequestHeaders.Add("User-Agent", 
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
        }

        public async Task<List<TournamentEntry>> ScrapeTabroomResults(string url, string tournamentName, string eventType)
        {
            try
            {
                Console.WriteLine($"Fetching data from: {url}");
                string html = await _httpClient.GetStringAsync(url);
                
                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                var entries = new List<TournamentEntry>();

                // Extract tournament year from the page (look for year in title or content)
                string year = ExtractYear(doc);

                // Debug: Count all tables on the page
                var allTables = doc.DocumentNode.Descendants("table").ToList();
                Console.WriteLine($"Found {allTables.Count} total tables on the page");

                // Determine event format by analyzing the page structure
                bool isCongress = eventType.ToUpper() == "CONGRESS" || eventType.ToUpper() == "CO";
                
                // Check if there are elimination rounds (Finals + Prelims sections)
                bool hasEliminationRounds = HasEliminationRounds(doc);

                if (isCongress)
                {
                    // Congress: Process separate Finals and Prelims tables
                    Console.WriteLine("Processing Congress event with separate Finals and Prelims tables...");
                    
                    var finalPlacesEntries = ProcessFinalPlacesTable(doc, tournamentName, eventType, year);
                    Console.WriteLine($"Finals entries found: {finalPlacesEntries.Count}");
                    entries.AddRange(finalPlacesEntries);

                    var prelimEntries = ProcessPrelimTable(doc, tournamentName, eventType, year);
                    Console.WriteLine($"Prelim entries found: {prelimEntries.Count}");
                    entries.AddRange(prelimEntries);
                }
                else if (hasEliminationRounds)
                {
                    // LD, PF, CX with elimination rounds: Process Finals + Prelims with sequential ranking
                    Console.WriteLine($"Processing {eventType} event with elimination rounds (Finals + Prelims)...");
                    var eliminationEntries = ProcessEliminationRounds(doc, tournamentName, eventType, year);
                    Console.WriteLine($"Elimination entries found: {eliminationEntries.Count}");
                    entries.AddRange(eliminationEntries);
                }
                else
                {
                    // LD, PF, CX, Classic: Process single table with implicit ranking by order
                    Console.WriteLine($"Processing {eventType} event with single ranked table...");
                    var rankedEntries = ProcessSingleRankedTable(doc, tournamentName, eventType, year);
                    Console.WriteLine($"Ranked entries found: {rankedEntries.Count}");
                    entries.AddRange(rankedEntries);
                }

                Console.WriteLine($"Successfully extracted {entries.Count} total entries from Tabroom.com");
                return entries;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error scraping Tabroom data: {ex.Message}");
                return new List<TournamentEntry>();
            }
        }

        private List<TournamentEntry> ProcessFinalPlacesTable(HtmlDocument doc, string tournamentName, string eventType, string year)
        {
            var entries = new List<TournamentEntry>();

            // Look for the Final Places table more specifically
            var finalPlacesTable = doc.DocumentNode
                .Descendants("table")
                .FirstOrDefault(table => 
                {
                    // Check if this table has ordinal place indicators (1st, 2nd, etc.)
                    var hasOrdinalPlaces = table.Descendants("td")
                        .Any(td => td.InnerText.Trim().EndsWith("st") || 
                                  td.InnerText.Trim().EndsWith("nd") || 
                                  td.InnerText.Trim().EndsWith("rd") || 
                                  td.InnerText.Trim().EndsWith("th"));
                    
                    // Make sure it's not the prelims table
                    var isNotPrelims = !table.OuterHtml.ToLower().Contains("prelim");
                    
                    return hasOrdinalPlaces && isNotPrelims;
                });

            if (finalPlacesTable != null)
            {
                Console.WriteLine("Found finals results table");
                var rows = finalPlacesTable.Descendants("tr").Skip(1); // Skip header row

                foreach (var row in rows)
                {
                    var cells = row.Descendants("td").ToList();
                    if (cells.Count >= 3)
                    {
                        var entry = ParseTableRow(cells, tournamentName, eventType, year, isPrelimsTable: false);
                        if (entry != null)
                        {
                            entries.Add(entry);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No finals results table found");
            }

            return entries;
        }

        private List<TournamentEntry> ProcessPrelimTable(HtmlDocument doc, string tournamentName, string eventType, string year)
        {
            var entries = new List<TournamentEntry>();

            // Debug: Look at all tables and their content
            var allTables = doc.DocumentNode.Descendants("table").ToList();
            for (int i = 0; i < allTables.Count; i++)
            {
                var table = allTables[i];
                var firstRow = table.Descendants("tr").Skip(1).FirstOrDefault();
                if (firstRow != null)
                {
                    var firstCell = firstRow.Descendants("td").FirstOrDefault();
                    Console.WriteLine($"Table {i}: First cell content = '{firstCell?.InnerText?.Trim()}'");
                }
            }

            // Look for the Prelim table - try multiple strategies
            var prelimTable = doc.DocumentNode
                .Descendants("table")
                .FirstOrDefault(table => 
                {
                    // Strategy 1: Check if table contains "Prelim" entries in the Place column
                    var hasPrelimEntries = table.Descendants("td")
                        .Any(td => td.InnerText.Trim().Equals("Prelim", StringComparison.OrdinalIgnoreCase));
                    
                    if (hasPrelimEntries)
                    {
                        Console.WriteLine("Found table with 'Prelim' entries");
                        return true;
                    }
                    
                    return false;
                });

            // If we didn't find it, try looking for the table that comes after "Protocol: Congress Prelims"
            if (prelimTable == null)
            {
                Console.WriteLine("Trying alternative prelim table detection...");
                var allNodes = doc.DocumentNode.Descendants().ToList();
                
                for (int i = 0; i < allNodes.Count; i++)
                {
                    var node = allNodes[i];
                    if (node.InnerText.Contains("Protocol: Congress Prelims") || 
                        node.InnerText.Contains("Congress Prelims"))
                    {
                        Console.WriteLine("Found 'Congress Prelims' text");
                        // Look for the next table after this text
                        for (int j = i + 1; j < allNodes.Count; j++)
                        {
                            if (allNodes[j].Name == "table")
                            {
                                prelimTable = allNodes[j];
                                Console.WriteLine("Found table after 'Congress Prelims' text");
                                break;
                            }
                        }
                        break;
                    }
                }
            }

            if (prelimTable != null)
            {
                Console.WriteLine("Found preliminary results table");
                var rows = prelimTable.Descendants("tr").Skip(1); // Skip header row
                int rowCount = 0;

                foreach (var row in rows)
                {
                    var cells = row.Descendants("td").ToList();
                    if (cells.Count >= 3)
                    {
                        rowCount++;
                        Console.WriteLine($"Processing prelim row {rowCount}: {cells[0]?.InnerText?.Trim()} | {cells[1]?.InnerText?.Trim()}");
                        
                        var entry = ParseTableRow(cells, tournamentName, eventType, year, isPrelimsTable: true);
                        if (entry != null)
                        {
                            entries.Add(entry);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No preliminary results table found");
            }

            return entries;
        }

        private bool HasEliminationRounds(HtmlDocument doc)
        {
            // Check if page has both "Final Places" and "Prelim" sections
            var pageText = doc.DocumentNode.InnerText.ToLower();
            bool hasFinalPlaces = pageText.Contains("final places");
            bool hasPrelims = pageText.Contains("prelim") && (pageText.Contains("debate prelims") || pageText.Contains("protocol:"));
            
            Console.WriteLine($"Has Final Places: {hasFinalPlaces}, Has Prelims: {hasPrelims}");
            return hasFinalPlaces && hasPrelims;
        }

        private List<TournamentEntry> ProcessEliminationRounds(HtmlDocument doc, string tournamentName, string eventType, string year)
        {
            var entries = new List<TournamentEntry>();

            // Process Finals table first (these get actual places 1, 2, 3, etc.)
            Console.WriteLine("Processing Finals section...");
            var finalsEntries = ProcessFinalPlacesTable(doc, tournamentName, eventType, year);
            entries.AddRange(finalsEntries);

            // Process Prelims table (these get sequential places starting after finals)
            Console.WriteLine("Processing Prelims section...");
            var prelimsEntries = ProcessEliminationPrelimsTable(doc, tournamentName, eventType, year, finalsEntries.Count);
            entries.AddRange(prelimsEntries);

            return entries;
        }

        private List<TournamentEntry> ProcessEliminationPrelimsTable(HtmlDocument doc, string tournamentName, string eventType, string year, int finalsCount)
        {
            var entries = new List<TournamentEntry>();

            // Find the prelims table (similar to existing logic but for non-Congress events)
            var prelimTable = doc.DocumentNode
                .Descendants("table")
                .FirstOrDefault(table => 
                {
                    // Look for table with "Prelim" entries and debate-specific columns
                    var hasPrelimEntries = table.Descendants("td")
                        .Any(td => td.InnerText.Trim().Equals("Prelim", StringComparison.OrdinalIgnoreCase));
                    
                    // Make sure it's not the Congress prelims table (which has different structure)
                    var hasDebateColumns = table.Descendants("th").Any(th => 
                        th.InnerText.Contains("WinPm") || th.InnerText.Contains("PtsPm"));
                    
                    return hasPrelimEntries && (hasDebateColumns || !eventType.ToUpper().Contains("CONGRESS"));
                });

            if (prelimTable != null)
            {
                Console.WriteLine("Found elimination prelims table");
                var rows = prelimTable.Descendants("tr").Skip(1).ToList(); // Skip header row
                
                for (int i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var cells = row.Descendants("td").ToList();
                    
                    if (cells.Count >= 3)
                    {
                        // Check if this row actually says "Prelim" in the place column
                        var placeColumn = cells[0].InnerText.Trim();
                        if (placeColumn.Equals("Prelim", StringComparison.OrdinalIgnoreCase))
                        {
                            // Assign sequential places starting after finals (3rd, 4th, 5th, etc.)
                            int sequentialPlace = finalsCount + i + 1;
                            
                            string entryName = cells[1].InnerText.Trim();
                            string school = cells[2].InnerText.Trim();
                            
                            Console.WriteLine($"Processing prelim row {sequentialPlace}: {entryName}");
                            
                            if (!string.IsNullOrEmpty(entryName) && !string.IsNullOrEmpty(school))
                            {
                                entries.Add(new TournamentEntry
                                {
                                    TournamentName = tournamentName,
                                    Year = year,
                                    Place = sequentialPlace,
                                    Entry = entryName,
                                    School = school,
                                    ElimPoints = 0 // No elimination points for prelim entries
                                });
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No elimination prelims table found");
            }

            return entries;
        }

        private List<TournamentEntry> ProcessSingleRankedTable(HtmlDocument doc, string tournamentName, string eventType, string year)
        {
            var entries = new List<TournamentEntry>();

            // Find the main results table - usually the first or largest table with student data
            var resultsTable = doc.DocumentNode
                .Descendants("table")
                .FirstOrDefault(table => 
                {
                    // Look for a table that has student names and schools
                    var hasStudentData = table.Descendants("td").Count() > 10; // Must have substantial data
                    var hasNameAndSchool = table.Descendants("tr").Skip(1).Take(3) // Check first few rows
                        .Any(row => row.Descendants("td").Count() >= 3); // At least 3 columns (place, name, school)
                    
                    return hasStudentData && hasNameAndSchool;
                });

            if (resultsTable != null)
            {
                Console.WriteLine("Found single ranked results table");
                var rows = resultsTable.Descendants("tr").Skip(1).ToList(); // Skip header row
                
                for (int i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var cells = row.Descendants("td").ToList();
                    
                    if (cells.Count >= 3)
                    {
                        // For single ranked tables, assign place based on row order (1st row = 1st place)
                        int implicitPlace = i + 1;
                        
                        Console.WriteLine($"Processing ranked row {implicitPlace}: {cells[1]?.InnerText?.Trim()}");
                        
                        var entry = ParseSingleRankedTableRow(cells, tournamentName, eventType, year, implicitPlace);
                        if (entry != null)
                        {
                            entries.Add(entry);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No single ranked results table found");
            }

            return entries;
        }

        private TournamentEntry? ParseSingleRankedTableRow(List<HtmlNode> cells, string tournamentName, string eventType, string year, int implicitPlace)
        {
            try
            {
                // For single ranked tables: Place is implicit, Entry is usually column 1, School is usually column 2
                string entryName = cells[1].InnerText.Trim();
                string school = cells[2].InnerText.Trim();
                
                // Skip empty or invalid rows
                if (string.IsNullOrEmpty(entryName) || string.IsNullOrEmpty(school))
                    return null;

                return new TournamentEntry
                {
                    TournamentName = tournamentName,
                    Year = year,
                    Place = implicitPlace,
                    Entry = entryName,
                    School = school,
                    ElimPoints = 0 // No elimination points for single table events
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing single ranked row: {ex.Message}");
                return null;
            }
        }

        private TournamentEntry? ParseTableRow(List<HtmlNode> cells, string tournamentName, string eventType, string year, bool isPrelimsTable)
        {
            try
            {
                string place;
                string entryName;
                string school;
                int elimPoints = 0;

                if (isPrelimsTable)
                {
                    // Prelim table format: Place | Entry | School | RkPm | RcpPm | Ballots
                    // Check if first column says "Prelim" - if so, this is the correct prelim table
                    var placeColumn = cells[0].InnerText.Trim();
                    if (placeColumn.Equals("Prelim", StringComparison.OrdinalIgnoreCase))
                    {
                        entryName = cells[1].InnerText.Trim();
                        school = cells[2].InnerText.Trim();
                        
                        // Use ranking from RkPm column (column 3) for place assignment
                        if (cells.Count > 3 && int.TryParse(cells[3].InnerText.Trim(), out int ranking))
                        {
                            place = (200 + ranking).ToString(); // Prelim places start at 201, 202, etc.
                        }
                        else
                        {
                            place = "999"; // Default high number for prelims without ranking
                        }
                    }
                    else
                    {
                        // This might be the wrong table, skip this row
                        return null;
                    }
                    elimPoints = 0; // No elimination points for prelims
                }
                else
                {
                    // Final Places table format: Place | Entry | School | ... | Ballots
                    place = ExtractPlaceNumber(cells[0].InnerText.Trim());
                    entryName = cells[1].InnerText.Trim();
                    school = cells[2].InnerText.Trim();
                    
                    // Try to extract elimination points from ballots or other columns
                    elimPoints = ExtractEliminationPoints(cells, int.Parse(place));
                }

                return new TournamentEntry
                {
                    TournamentName = tournamentName,
                    Year = year,
                    Place = int.Parse(place),
                    Entry = entryName,
                    School = school,
                    ElimPoints = elimPoints
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing row: {ex.Message}");
                return null;
            }
        }

        private string ExtractPlaceNumber(string placeText)
        {
            // Extract number from "1st", "2nd", "3rd", etc.
            var numbers = new string(placeText.Where(char.IsDigit).ToArray());
            return string.IsNullOrEmpty(numbers) ? "999" : numbers;
        }

        private int ExtractEliminationPoints(List<HtmlNode> cells, int place)
        {
            // For now, return 0 - we'll handle all scoring in the main program
            // This prevents double-counting of points
            return 0;
        }

        private string ExtractYear(HtmlDocument doc)
        {
            // Look for year in the tournament title or page content
            var titleNode = doc.DocumentNode.SelectSingleNode("//title");
            if (titleNode != null)
            {
                var titleText = titleNode.InnerText;
                var yearMatch = System.Text.RegularExpressions.Regex.Match(titleText, @"20\d{2}");
                if (yearMatch.Success)
                {
                    return yearMatch.Value;
                }
            }

            // Look for year in tournament header
            var headerNodes = doc.DocumentNode.SelectNodes("//h1 | //h2 | //h3 | //h4 | //h5");
            if (headerNodes != null)
            {
                foreach (var node in headerNodes)
                {
                    var yearMatch = System.Text.RegularExpressions.Regex.Match(node.InnerText, @"20\d{2}");
                    if (yearMatch.Success)
                    {
                        return yearMatch.Value;
                    }
                }
            }

            // Default to current year if not found
            return DateTime.Now.Year.ToString();
        }

        public async Task AppendToEventCSV(List<TournamentEntry> entries, string eventType)
        {
            try
            {
                string fileName = GetEventFileName(eventType);
                bool fileExists = File.Exists(fileName);
                
                using (var writer = new StreamWriter(fileName, append: true))
                {
                    // Write header only if file doesn't exist
                    if (!fileExists)
                    {
                        await writer.WriteLineAsync("Tournament Name,Year,Place,Entry,School,ElimPoints");
                    }

                    foreach (var entry in entries)
                    {
                        var line = $"{entry.TournamentName},{entry.Year},{entry.Place},{entry.Entry},{entry.School},{entry.ElimPoints}";
                        await writer.WriteLineAsync(line);
                    }
                }

                Console.WriteLine($"Data appended to {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving to CSV: {ex.Message}");
            }
        }

        private string GetEventFileName(string eventType)
        {
            return eventType.ToUpper() switch
            {
                "CONGRESS" or "CO" => "CODebate.csv",
                "LD" or "LINCOLN-DOUGLAS" => "LDDebate.csv", 
                "PF" or "PUBLIC FORUM" => "PFDebate.csv",
                "CX" or "POLICY" or "CROSS-EXAMINATION" => "CXDebate.csv",
                "CLASSIC" => "ClassicDebate.csv",
                _ => $"{eventType}Debate.csv"
            };
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    public class TournamentEntry
    {
        public string TournamentName { get; set; } = string.Empty;
        public string Year { get; set; } = string.Empty;
        public int Place { get; set; }
        public string Entry { get; set; } = string.Empty;
        public string School { get; set; } = string.Empty;
        public int ElimPoints { get; set; }
    }
}
