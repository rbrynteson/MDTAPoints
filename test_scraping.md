# Testing Tabroom.com Web Scraping

## Sample Test Case

**URL**: 
https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=37303&result_id=401967 (CO)
https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=37303&result_id=401972 (LD)

**Tournament Name**: Coon Rapids Cardinal Debate Tournament

**Event Type**: Congress

## Expected Results

The scraper should extract:

### Final Places Table
- 12 entries from "Final Places" section
- Places 1-12 with student names, schools, and calculated elimination points

### Prelim Table  
- 19 entries from "Prelim" section
- All marked as place 999 (prelim indicator)
- Elimination points calculated from ranking

## Test Command

When running the application:
1. Choose option "2" (Tabroom.com URL)
2. Enter URL: `https://www.tabroom.com/index/tourn/results/event_results.mhtml?tourn_id=37303&result_id=401967`
3. Enter tournament name: `Coon Rapids Cardinal Debate Tournament`
4. Enter event type: `Congress`

## Expected Output

- CSV file: `Coon_Rapids_Cardinal_Debate_Tournament_Congress.csv`
- 31 total entries (12 final + 19 prelim)
- Proper MDTA scoring calculations applied
