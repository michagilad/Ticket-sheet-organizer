# Ticket Sheet Organizer

A Python application that processes CSV exports of tickets from experience management systems and creates organized XLSX output with intelligent grouping and formatting.

## Features

- **Robust CSV Reading**: Handles line breaks, special characters, and various CSV formats
- **Column Selection & Reordering**: Extracts only the required columns and reorders them for optimal viewing
- **Experience Grouping**: Automatically groups tickets by experience with visual indicators
- **Smart Cell Merging**: Merges cells for experiences with multiple tickets to clearly show which tickets belong together
- **Alternating Colors**: Uses color coding to distinguish between different experiences
- **Professional Formatting**: Clean, readable output with proper borders, alignment, and column widths

## Installation

1. Ensure you have Python 3.7 or higher installed
2. Install the required dependencies:

```bash
pip3 install -r requirements.txt
```

## Usage

You can use this application in two ways: through a **Web UI** (recommended) or via **Command Line**.

### Option 1: Web UI (Recommended)

1. Start the web server:

```bash
python3 app.py
```

2. Open your browser and go to: `http://localhost:5001`

3. Drag and drop your CSV file or click to browse

4. Click "Process & Download XLSX" to get your organized file

The web UI features:
- Beautiful drag-and-drop interface
- Instant file processing
- Automatic download of the organized XLSX file
- No need to remember command line syntax

### Option 2: Command Line

#### Basic Usage

```bash
python3 ticket_organizer.py input_tickets.csv
```

This will create an output file named `input_tickets_organized.xlsx` in the same directory.

#### Specify Output File

```bash
python3 ticket_organizer.py input_tickets.csv output_name.xlsx
```

## Input Format

The input CSV should contain ticket export data with the following columns:

- Ticket ID
- Backstage Experience page
- Company
- Create date
- Priority
- Record source detail 1
- Ticket description
- Ticket name
- Ticket status
- Associated Experience IDs
- Associated Experience
- Assignee
- Valid ticket
- Ticket Category
- Sub Category
- Non-ticketed issues
- Additional Notes
- Notes - Admin
- Next step
- Last Updated - Status
- Last Updated - Next step

**Note**: The CSV can contain additional columns, but only the above columns will be processed.

## Output Format

The XLSX output will have columns in the following order:

1. Associated Experience (bold text)
2. Backstage Experience page (clickable URLs)
3. Associated Experience IDs
4. Assignee
5. Ticket name (bold text)
6. Ticket description
7. Valid ticket
8. Ticket Category
9. Ticket status (with dropdown: Todo, In Progress, Done, Resolved, Stuck)
10. Sub Category
11. Non-ticketed issues
12. Additional Notes
13. Notes - Admin
14. Next step
15. Create date
16. Ticket ID
17. Record source detail 1
18. Company
19. Priority
20. Last Updated - Status
21. Last Updated - Next step

### Visual Organization

- **Merged Cells**: When an experience has multiple tickets, the first four columns (Associated Experience, Backstage Experience page, Associated Experience IDs, and Assignee) are merged vertically
- **Bold Text**: Associated Experience and Ticket name columns are displayed in bold for easier scanning
- **Clickable URLs**: Backstage Experience page URLs are automatically converted to clickable hyperlinks (blue, underlined)
- **Color Coding**: Each experience group has alternating background colors (gray and light blue) to make it easy to distinguish between different experiences
- **Borders**: Merged cells have thicker borders to clearly delineate experience groups
- **Frozen Header**: The header row is frozen for easy scrolling through large datasets

## Example

If your input CSV has:
- Experience A with 3 tickets
- Experience B with 1 ticket
- Experience C with 2 tickets

The output will show:
- 3 rows with merged experience identification cells (colored gray), each showing a different ticket
- 1 row for Experience B (colored light blue)
- 2 rows with merged experience identification cells (colored gray), each showing a different ticket

This makes it immediately clear which tickets belong to the same experience while still displaying all individual ticket details.

## Troubleshooting

### Missing Columns Error

If you get an error about missing columns, ensure your input CSV contains all the required column headers listed in the "Input Format" section.

### Character Encoding Issues

The script automatically handles UTF-8 with BOM encoding. If you encounter encoding issues, try opening your CSV in Excel or Google Sheets and re-exporting it as UTF-8.

### Line Break Issues

The CSV reader is configured to handle line breaks within fields properly. Make sure your CSV uses proper quoting for fields that contain line breaks.

## Requirements

- Python 3.7+
- openpyxl 3.1.2 or higher
- Flask 3.0.0 or higher (for web UI)

## License

This tool is provided as-is for organizing ticket export data.


