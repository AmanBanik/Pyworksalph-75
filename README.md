# 📊 NEET OMR Scoring Software (Excel-Based)

A Python program that evaluates NEET-style OMR responses using only Excel files and pandas. Designed for offline, fast, and simple use — ideal for students and educators who prefer using spreadsheet inputs instead of scanned sheets.

---

## 🧠 Project Objective

This tool helps in:
- Calculating scores from manually filled NEET OMR responses
- Comparing user-marked answers to an official answer key
- Handling unattempted and invalid answers gracefully
- Exporting detailed results to a new Excel file (
eet_scoring_details.xlsx)

---

## 📂 How It Works

1. You fill out an Excel sheet of **your answers**.
2. You also provide an Excel **answer key**.
3. The program reads both files and:
   - Matches answers
   - Applies NEET marking scheme (+4 for correct, -1 for wrong, 0 for unattempted)
   - Saves detailed question-wise analysis to a new Excel sheet

---
