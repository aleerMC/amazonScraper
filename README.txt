Amazon Top-20 → Micro Center Matcher (no pyarrow)

1) Double-click launch_app.bat
   - First run: creates a local venv and installs requirements
   - Opens http://localhost:8501 in your browser

2) Usage
   - Paste an Amazon Best Sellers URL and a short description
   - Click "Fetch Top 20"
   - Use the right column to search Micro Center by SKU/keywords
   - Pick a candidate and "Submit to Micro Center column"
   - Add Attribute Match + Notes
   - Save as New (left sidebar keeps your saved searches)
   - Download Spreadsheet (.xlsx) for PowerPoint screenshots

3) Saved data
   - Stored as CSV under .saved_searches/<id>/data.csv
   - Metadata at .saved_searches/<id>/meta.json
   - No pyarrow required
