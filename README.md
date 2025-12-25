# Ticket-Resolve-Tool

This tool helps in organizing PDF files based on ticket status from an Excel sheet.

## Requirements

- Python 3.x
- Install dependencies: `pip install -r requirements.txt`

## Usage

Run the script `move_pdfs.py` to open the GUI.

```bash
python move_pdfs.py
```

The GUI will appear with fields for:
- Excel File Path (use Browse button to select)
- PDF Folder Path (folder containing the PDF files)
- Destination Folder Path (where to move "Replied" and "Forwarded" files)
- Ticket Column Name (default: "Ticket Number")
- Status Column Name (default: "Status")

Fill in the fields and click "Run". A message will show the number of files moved or any errors.

## Notes

- Ensure the Excel file has the correct column names. If different, use the optional arguments.
- The script assumes ticket numbers are strings and the last part of the filename before .pdf is the ticket number.
- Files are moved, not copied.