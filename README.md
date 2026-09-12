# 🎓 Student Result Management System

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![openpyxl](https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Status](https://img.shields.io/badge/status-active-success.svg)]()

A menu-driven **command-line Python application** for managing student academic results — enter student details and marks, get automatic total/percentage/grade/pass-fail calculation, and keep every record permanently in an Excel workbook via `openpyxl`.

---

## 📑 Table of Contents

- [About the Project](#-about-the-project)
- [Key Features](#-key-features)
- [Demo](#️-demo)
- [Grading System](#-grading-system)
- [Excel Data Structure](#-excel-data-structure)
- [Project Structure](#-project-structure)
- [Working Flow](#️-working-flow)
- [Getting Started](#-getting-started)
- [Python Concepts Used](#-python-concepts-used)
- [Future Enhancements](#-future-enhancements)
- [License](#-license)

---

## 📌 About the Project

The **Student Result Management System** is a lightweight, dependency-minimal CLI tool built to simplify basic student result management for a class or course.

The application collects, per student:
- Roll Number
- Name
- Class / Course
- Marks for 5 subjects

...and automatically computes:
- ✅ Total Marks
- ✅ Percentage
- ✅ Grade
- ✅ Pass / Fail status

Every record is saved permanently to `student_results.xlsx`, so data survives between runs — no database setup required.

## ✨ Key Features

| Category | Details |
|---|---|
| 👨‍🎓 **Student Management** | Add new results, prevent duplicate roll numbers, validate all inputs |
| 📊 **Result Calculation** | Auto-computes total, percentage, grade, and pass/fail status |
| 🔍 **Result Search** | Look up any student's full result instantly by roll number |
| 📋 **Bulk View** | Display every stored record in a clean, aligned table |
| 📁 **Excel Storage** | Auto-creates the workbook on first run; reads/writes via `openpyxl` |
| 🛡️ **Input Validation** | Rejects out-of-range marks, non-numeric input, and duplicate roll numbers |

## 🖥️ Demo

```
========================================
      STUDENT RESULT MANAGEMENT
========================================
1. Add Student Result
2. Get Student Result
3. Show All Student Data
4. Exit

Enter your choice: 2
Enter Roll No to search: 101

--------------------------------
Student Result
--------------------------------
Roll No     : 101
Name        : Rahul
Class       : BCA
Total       : 425
Percentage  : 85.00%
Grade       : A
Status      : PASS
--------------------------------
```

**Show All Student Data**

```
Roll No   Name           Class     Total     Percentage   Grade   Status
--------------------------------------------------------------------------
101       Rahul          BCA       425       85.00        A       PASS
102       Priya          BCA       378       75.60        B       PASS
103       Amit           BCA       198       39.60        F       FAIL
```

## 🧮 Grading System

| Percentage | Grade |
|---|---|
| 90% and above | A+ |
| 80% – 89% | A |
| 70% – 79% | B |
| 60% – 69% | C |
| 40% – 59% | D |
| Below 40% | F |

**Pass / Fail rule:** a student fails if their overall percentage is below 40%, **or** if any single subject score is below 33 — matching the common Indian university passing convention.

## 📊 Excel Data Structure

Records are stored in **`student_results.xlsx`**, with the following columns:

| Column | Description |
|---|---|
| Roll No | Unique student roll number |
| Name | Student name |
| Class | Class / Course |
| Subject 1–5 | Marks obtained in each subject |
| Total | Sum of all subject marks |
| Percentage | Total ÷ number of subjects |
| Grade | Calculated from percentage |
| Status | PASS / FAIL |

## 📁 Project Structure

```
.
├── student_result_management.py   # Main application
├── student_results.xlsx           # Auto-created on first run (gitignored)
├── requirements.txt                # Project dependencies
├── .gitignore
└── README.md
```

## 🏗️ Working Flow

```
                ┌──────────────────────────┐
                │           START          │
                └────────────┬─────────────┘
                             ↓
                ┌──────────────────────────┐
                │  Create/Check Excel File │
                │   student_results.xlsx   │
                └────────────┬─────────────┘
                             ↓
                ┌──────────────────────────┐
                │         MAIN MENU        │
                ├──────────────────────────┤
                │ 1. Add Student Result    │
                │ 2. Get Student Result    │
                │ 3. Show All Student Data │
                │ 4. Exit                  │
                └────────────┬─────────────┘
                             ↓
        ┌────────────────────┼────────────────────┐
        ↓                    ↓                     ↓
  ┌───────────┐       ┌─────────────┐       ┌─────────────┐
  │  Option 1 │       │  Option 2   │       │  Option 3   │
  │  Add Data │       │  Get Data   │       │  Show All   │
  └─────┬─────┘       └──────┬──────┘       └──────┬──────┘
        ↓                    ↓                     ↓
  Enter details        Enter Roll No.        Read Excel data
        ↓                    ↓                     ↓
  Enter 5 subjects     Search record          Print table
        ↓                    ↓
  Calculate Total,     Display result
  % & Grade
        ↓
  Determine PASS/FAIL
        ↓
  Save to Excel
        └────────────────────┬────────────────────┘
                             ↓
                       Return to Menu
                             ↓
                     Option 4 chosen?
                             ↓
                            EXIT
```

## 🚀 Getting Started

### Prerequisites
- Python 3.7+
- pip

### Installation

```bash
git clone https://github.com/namratabaviskar18-rgb/student-result-management.git
cd student-result-management
pip install -r requirements.txt
```

### Run

```bash
python student_result_management.py
```

The program automatically creates `student_results.xlsx` in the project folder the first time you run it.

## 📚 Python Concepts Used

This project demonstrates practical use of:

- Variables & data types
- `input()` and formatted output
- Lists
- Functions, parameters & return values
- `if-elif-else` branching
- `for` loops (subject entry)
- `while` loop (menu system)
- `try-except` exception handling
- String formatting (f-strings, alignment)
- File handling & Excel I/O with `openpyxl`
- Input validation & duplicate-record checks

## 🔮 Future Enhancements

- ✏️ Edit or delete an existing student record
- 🔎 Search by student name in addition to roll number
- 🏆 Class-wise topper list and basic analytics
- 📄 Export individual results as PDF report cards
- 🌐 Optional GUI or web front-end

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

### 👩‍💻 Author

Built as a hands-on Python + Excel automation project — contributions and suggestions are welcome!
