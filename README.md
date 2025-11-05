# MDTAPoints

A C# console application for calculating and analyzing tournament points for Minnesota Debate Tournament Association (MDTA) competitions. This tool processes Excel files containing tournament results and generates comprehensive scoring reports for both individual students and schools.

## Features

- **Excel File Processing**: Reads tournament data from Excel (.xlsx) files using EPPlus library
- **Web Scraping**: Automatically extracts tournament data from Tabroom.com results pages
- **Individual Student Scoring**: Calculates points for each student participant across multiple tournaments
- **School Team Scoring**: Aggregates top 2 performers per school for each tournament
- **Flexible Bonus System**: Awards bonus points based on tournament size and placement
- **Comprehensive Reports**: Generates detailed CSV-style output for analysis
- **Dual Input Methods**: Support for both manual Excel files and automated web scraping

## Scoring System

### Base Points
- **Participation**: 1 point for participating in any tournament
- **Elimination Points**: Additional points based on elimination rounds (read from Excel column 9)

### Bonus Points
The application uses a tiered bonus system:

#### Small Tournaments (≤12 entries)
- Top half of participants receive 1 bonus point

#### Large Tournaments (>12 entries)
- 1st place: 2 bonus points
- 2nd-8th place: 1 bonus point

### School Scoring
- Takes the top 2 individual scores from each school per tournament
- Sums these across all tournaments for final school rankings

## Prerequisites

- .NET 8.0 or later
- Windows, macOS, or Linux

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd MDTAPoints
```

2. Build the project:
```bash
dotnet build
```

## Usage

1. Run the application:
```bash
dotnet run --project MDTAPoints
```

2. Choose your input method:
   - **Option 1**: Excel file (traditional method)
   - **Option 2**: Tabroom.com URL (web scraping)

### Option 1: Excel File Processing
- Enter the full path to your Excel file containing tournament data
- File must follow the specified column format (see below)

### Option 2: Web Scraping from Tabroom.com
- Enter the Tabroom.com results URL (e.g., `https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=37303&result_id=401967`)
- Enter tournament name (e.g., "Coon Rapids Cardinal Debate Tournament")
- Enter event type (e.g., "Congress", "LD", "PF")
- The application will automatically scrape the data and create a CSV file

3. The application will process the data and display:
   - Individual student scores across all tournaments
   - School team scores (top 2 per school per tournament)

## Excel File Format

The application expects Excel files with the following column structure:

| Column | Description |
|--------|-------------|
| A (1) | Tournament Name |
| B (2) | *Not used* |
| C (3) | Year |
| D (4) | Place/Ranking |
| E (5) | Entry (Student names, separated by '&' for partnerships) |
| F (6) | School |
| G-H (7-8) | *Not used* |
| I (9) | Elimination Points |

### Sample Data Format
```
Tournament Name | | 2024 | 1 | John Doe & Jane Smith | Lincoln High | | | 3
State Championship | | 2024 | 2 | Bob Johnson | Washington Prep | | | 2
```

## Supported Tournament Types

Based on the project configuration, this tool is designed to work with:
- Classic Debate tournaments
- Congressional Debate (CO) tournaments  
- Lincoln-Douglas (LD) Debate tournaments

## Tabroom.com Integration

The application can automatically scrape tournament results from [Tabroom.com](https://www.tabroom.com), the official tournament management platform used by the National Speech and Debate Association.

### Supported Tabroom.com Features
- **Final Places Tables**: Extracts ranked results with placements
- **Preliminary Results**: Processes prelim rounds data
- **Multiple Event Types**: Congress, Lincoln-Douglas, Public Forum, etc.
- **Automatic Data Conversion**: Converts Tabroom format to MDTA scoring system

### How to Find Tabroom URLs
1. Go to [Tabroom.com](https://www.tabroom.com)
2. Navigate to Results → Find your tournament
3. Click on the specific event results
4. Copy the URL from your browser (should contain `event_results.mhtml`)

### Example URLs
- Congress: `https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=37303&result_id=401967`
- LD Debate: `https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=12345&result_id=67890`

## Output Format

### Student Scores
```
Entry, School, Tournaments, [Tournament1], [Tournament2], ..., Total Points
Lincoln High:John Doe, Lincoln High, 2, 4, 3, 7
Washington Prep:Bob Johnson, Washington Prep, 1, 0, 3, 3
```

### School Scores
```
School, [Tournament1], [Tournament2], ..., Total Points
Lincoln High, 4, 3, 7
Washington Prep, 0, 3, 3
```

## Dependencies

- **EPPlus (v7.5.1)**: Excel file processing library
- **HtmlAgilityPack (v1.11.54)**: HTML parsing for web scraping
- **System.Net.Http (v4.3.4)**: HTTP client for web requests
- **.NET 8.0**: Runtime framework

## License

This project uses EPPlus under a non-commercial license context. Ensure compliance with EPPlus licensing terms for your use case.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Support

For issues or questions regarding tournament scoring rules, please contact the Maryland Debate Tournament Association.

## Technical Notes

- The application uses top-level statements (C# 9.0+ feature)
- Memory-efficient processing using LINQ and grouping operations
- Handles partnership entries by splitting names on '&' character
- Robust error handling for missing or invalid data