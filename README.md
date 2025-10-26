# Compare Two CSV Files
A simple Python GUI application to compare and report changes between old and new csv file versions. For a proper demo, visit the **"Releases"** tab and open the "demo_app.zip" file (there is both a Windows and Mac version).

---
## Report Generation Process

1. **Browse and Select Two CSV Files**: one representing the "Before" state, the other the "After" state.
2. **Select Key Column**: The key column uniquely identifies rows in both CSV (Only columns common to both CSVs are selectable as the key)
3. **Compare CSV Files**: Identifies added rows, removed rows, and changed cells.
4. **Preview Results**: Preview report of comparison result
5. **Export Report**: Export the comparison summary to a Word document (.docx). Includes the following: summary counts, preview of changed values (first 200), and added and removed rows (based on key column)
