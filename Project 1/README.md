<div align="center">

# -- ! Formula Forge ! --
### *A Multi-Domain Excel Analytics Workbook — Student, Sales & Employee Records*

[![Excel](https://img.shields.io/badge/Excel-2021%2F365-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-Logical%20%2B%20Lookup%20%2B%20Text-00758F?style=for-the-badge&logo=databricks&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![Sheets](https://img.shields.io/badge/Sheets-3%20Data%20Domains-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![Dynamic Arrays](https://img.shields.io/badge/Dynamic%20Arrays-FILTER%20%2F%20XLOOKUP-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)

<br/>

> *"A raw column of numbers is data; a formula next to it is an opinion about what matters."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Workbook Structure](#️-workbook-structure)
- [🎓 Sheet 1 — Student](#-sheet-1--student)
- [📊 Sheet 2 — Sales Data](#-sheet-2--sales-data)
- [👔 Sheet 3 — Employee Data](#-sheet-3--employee-data)
- [🧠 Formula Library](#-formula-library)
- [⚠️ Notes & Data Consistency](#️-notes--data-consistency)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)

---

## 📌 Overview

**Formula Forge** is a pure-Excel practical spanning three unrelated data domains in one workbook: student academic records, regional sales transactions, and employee HR data. Each sheet is a self-contained formula showcase — no VBA, no Power Query, just cell formulas doing the work of grading, discounting, ranking, and looking up records.

This project is designed to:
- Practice **logical branching** with nested `IF`, `IFS`, and `AND`/`OR`
- Apply **lookup functions** — `VLOOKUP`, `XLOOKUP`, `XMATCH`, and `INDEX`/`MATCH` (including an array-form multi-criteria match)
- Use **dynamic array functions** (`FILTER`) that spill results across multiple cells from one formula
- Work with **text functions** (`LEFT`, `FIND`, `UPPER`, `LOWER`) to parse and reformat names
- Work with **date functions** (`DATEDIF`, `TODAY`) to compute live ages
- Use **math/rounding functions** (`ROUND`, `CEILING.MATH`, `FLOOR.MATH`)
- Use **aggregate and reference functions** (`COUNTIF`, `SUMIFS`, `OFFSET`, `INDIRECT`, `SUM`)

---

## 🎯 Problem Statement

> **Objective:** Given three independent flat datasets — students, sales transactions, and employees — build formula-driven columns and lookup panels that answer common reporting questions without touching the source data.

Each sheet poses its own mini-problem: grade and rank students from raw subject scores, calculate discount eligibility and totals from a sales log, and build quick employee lookup panels from an HR roster. All three are solved with formulas alone — the source rows are never manually recalculated.

| 📂 Sheet | 📄 Role | 🔍 Description |
|----------|---------|-----------------|
| `Student` | Academic records | 10 students, subject scores, grading and ranking formulas |
| `Sales Data` | Transaction log | 10 sales rows across 5 regions, discount and totals formulas |
| `Employee Data` | HR roster | 10 employees, salary rounding and ID-based lookup panels |

The goal is to demonstrate **formula fluency across independent domains** — the same handful of Excel function families, reapplied to very different data.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧮 **Nested Conditional Grading** | `IF` + `IFS` combined to assign letter grades, with a hard override for failing scores |
| 🔀 **Multi-Condition Flags** | `AND`/`OR` inside `IF` for pass/fail and discount-eligibility flags |
| 💧 **Dynamic Array Filtering** | `FILTER()` spills a live list of students scoring 80+ across two columns |
| 🔍 **Four Lookup Styles** | `VLOOKUP`, `XLOOKUP`, `XMATCH`, and array-form `INDEX`/`MATCH` (including multi-criteria) |
| 🧵 **Text Parsing** | `LEFT` + `FIND` extract first names; `UPPER`/`LOWER` reformat casing |
| 📅 **Live Age Calculation** | `DATEDIF(dob, TODAY(), "y")` — recalculates automatically as the file ages |
| 📊 **Conditional Aggregation** | `COUNTIF` and `SUMIFS` for threshold counts and grouped totals |
| 🧭 **Indirect Referencing** | `OFFSET` for a sliding "recent totals" window; `INDIRECT` to sum a range named as text |
| 🔢 **Custom Rounding** | `ROUND`, `CEILING.MATH`, `FLOOR.MATH` rounding salaries to the nearest 100 |

---

## 🏗️ Workbook Structure

```
📦 formula-forge/
│
└── 📄 Emp_Management_Ved.xlsx
    │
    ├── 📄 Student          ← Grades, ranking, name parsing, VLOOKUP panel
    ├── 📄 Sales Data       ← Discounts, SUMIFS, XMATCH, OFFSET, INDIRECT
    └── 📄 Employee Data    ← Salary rounding, INDEX/MATCH & XLOOKUP panels
```

**Columns per sheet:**

```
Student                          Sales Data                    Employee Data
────────                          ───────────                    ─────────────
Student ID                        Sales ID                        Employee ID
Full Name                         Region                          Name
DOB                                Product                         Department
Math                               Product code                    Salary
Science                            Sales Person                    Joining Date
Average (formula)                  Month                           Salary (Round)
Grade (formula)                    Price                           Salary (Ceil)
Above 80 in SCI & MATH (formula)   Quantity                        Salary (Floor)
Average Score Above 60 (formula)   Sales Amount (formula)
First Name (formula)               Date
Name (Upper) / (Lower) (formula)   Discount % (formula)
Age (formula)                      Discount Eligible (formula)
```

---

## 🎓 Sheet 1 — Student

10 students with Math and Science scores, graded and ranked entirely through formulas.

**Key Formulas:**
```excel
Average (F2):        =AVERAGE(D2:E2)
Grade (G2):           =IF(OR(D2<35, E2<35), "F",
                          IFS(AVERAGE(D2:E2)>90,"A", AVERAGE(D2:E2)>75,"B",
                              AVERAGE(D2:E2)>60,"C", AVERAGE(D2:E2)>50,"D", TRUE,"F"))
Above 80 both (H2):   =IF(AND(D2>80, E2>80), "YES", "NO")
Avg if both>60 (I2):  =IF(AND(D2>60, E2>60), AVERAGE(D2:E2), "-")
First Name (J2):      =LEFT(B2, FIND(" ", B2)-1)
Name Upper/Lower:      =UPPER(J2)   /   =LOWER(J2)
Age (M2):              =DATEDIF(C2, TODAY(), "y")
Count Math>50 (Q1):    =COUNTIF(D2:D11, ">50")
VLOOKUP panel (Q5):    =VLOOKUP(Q3, A2:H11, Q4, FALSE)   ← ID and column number set in Q3/Q4
Top scorers (B15):    =FILTER(B2:B11, F2:F11>=80)        ← spills names + scores of the 80+ crowd
```

**Sample Computed Output:**

| Student | Avg | Grade | Above 80 in Both | First Name |
|---------|-----|-------|-------------------|------------|
| Kabir Mehta (88, 94) | 91.0 | A | YES | Kabir |
| Ananya Sharma (91, 89) | 90.0 | B | YES | Ananya |
| Dev Patel (72, 65) | 68.5 | C | NO | Dev |
| Isha Verma (95, 98) | 96.5 | A | YES | Isha |
| Rohan Kapoor (81, 79) | 80.0 | B | NO | Rohan |

*(All 10 rows follow the same pattern — every student here clears the ">60 in both subjects" threshold, so the "Average Score Above 60" column mirrors the plain Average column for every row.)*

The `COUNTIF(D2:D11, ">50")` count comes out to **10** — every student scored above 50 in Math. The VLOOKUP panel (`Q3 = Student ID 10`, `Q4 = Column 2`) fetches **Riya Sen**, the Full Name in column B for student 10.

---

## 📊 Sheet 2 — Sales Data

10 transactions across 5 regions and 4 products, with discount logic and several independent lookup/aggregation panels.

**Key Formulas:**
```excel
Sales Amount (I2):     =G2*H2                                             (Price × Quantity)
Discount % (K2):        =IFS(G2>=50000,"15%", G2>=25000,"10%",
                             G2>=10000,"5%", G2>=5000,"2%", TRUE,"0%")     (tiered by Price)
Discount Eligible (L2): =IF(OR(G2>=5000, H2>=50), "YES", "NO")
Region+Product total:   =SUMIFS(I2:I11, B2:B11, A15, C2:C11, B15)
Multi-criteria match:   =INDEX(I2:I11, MATCH(1, (E2:E11=E15)*(F2:F11=F15), 0))
Column position:         =XMATCH(I15, C2:C11, 0, 1)
Price by product code:  =VLOOKUP(A18, D2:G11, 4, FALSE)
Person's sales lookup:  =XLOOKUP(E18, E2:E11, I2:I11, "-", 0, 1)
Sliding window total:   =SUM(OFFSET(I11, -(M15-1), 0, M15, 1))
Sum via text range:      =SUM(INDIRECT(I18))
```

**Sample Computed Output:**

| Metric | Result |
|--------|--------|
| Total sales amount, all 10 rows (`SUM(INDIRECT("I2:I11"))`) | **₹17,11,000** |
| East + Keyboard total (`SUMIFS`) | **₹90,000** (matches the single Feb East/Keyboard row) |
| Discount-ineligible row | Only **Kavya Singhania — Keyboard, May** (₹1,800 price, 35 units — under both thresholds) |
| Column position of "Keyboard" via `XMATCH` | **4th** position in the Product column |
| Price for product code 101 via `VLOOKUP` | **₹60,000** (Laptop) |

---

## 👔 Sheet 3 — Employee Data

10 employees with salary rounding variants and two independent ID-driven lookup panels.

**Key Formulas:**
```excel
Round to nearest whole:  =ROUND(D2, 0)
Round up to nearest 100:  =CEILING.MATH(D2, 100)
Round down to nearest 100: =FLOOR.MATH(D2, 100)
Name by ID (INDEX/MATCH):  =INDEX(B2:B11, MATCH($B$13, A2:A11, 0))
Salary by ID (XLOOKUP):     =XLOOKUP(E13, A2:A11, D2:D11, "-", 0, 1)
```

**Sample Computed Output:**

| Employee ID | Name | Department | Salary | Round / Ceil / Floor |
|-------------|------|------------|--------|------------------------|
| 1 | Manish Agarwal | Sales | ₹67,000 | 67,000 / 67,000 / 67,000 |
| 5 | Arjun Fernandez | IT | ₹95,000 | 95,000 / 95,000 / 95,000 |
| 8 | Simran Kaur | Finance | ₹98,000 | 98,000 / 98,000 / 98,000 |

Both lookup panels (`B13`/`E13` set to Employee ID `1`) resolve to **Manish Agarwal, Sales, ₹67,000** — one panel via `INDEX`/`MATCH`, the other via `XLOOKUP`, demonstrating two ways to solve the same lookup.

---

## 🧠 Formula Library

Grouped by function family, across all three sheets.

### 🔀 Logical
`IF`, `IFS`, `AND`, `OR` — grading, discount eligibility, pass/fail flags

### 🔍 Lookup & Reference
`VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`/`MATCH` (including array multi-criteria form), `OFFSET`, `INDIRECT`

### 💧 Dynamic Arrays
`FILTER` — spills a live, auto-updating list of top-scoring students

### 🧵 Text
`LEFT`, `FIND`, `UPPER`, `LOWER`

### 📅 Date
`DATEDIF`, `TODAY`

### 🔢 Math & Rounding
`ROUND`, `CEILING.MATH`, `FLOOR.MATH`

### 📊 Aggregation
`AVERAGE`, `COUNTIF`, `SUMIFS`, `SUM`

---

## ⚠️ Notes & Data Consistency

A few things worth knowing before treating every cell as authoritative:

- **Age column is live, not static.** `DATEDIF(..., TODAY(), "y")` recalculates every time the file is opened, so the ages shown will drift forward over time — they aren't a fixed dataset value.
- **The "Top Scorers" panel (rows 16–18) has hardcoded values that don't match the `FILTER` formula's actual source data.** `FILTER(B2:B11, F2:F11>=80)` should spill names from the Student list above (Kabir, Ananya, Isha, Rohan, Tara, Vihaan, Riya), but rows 16–18 contain **Sneha Iyer, Diya Patel, and Vivaan Gupta** — none of whom appear anywhere in the Student table. These look like leftover manual values from an earlier version of the sheet, sitting inside what should be the formula's spill range.
- **Two Sales Data lookups reference people who aren't in the dataset.** The multi-criteria `INDEX`/`MATCH` panel looks up **"Rahul Desai"** and the `XLOOKUP` panel looks up **"Sonal Kapoor"** — neither appears in the `Sales Person` column (which only contains Vikram Malhotra, Pooja Hegde, Siddharth Roy, and Kavya Singhania). The `XLOOKUP` panel has a `"-"` fallback so it degrades gracefully; the `MATCH`-based panel has no such fallback and will return `#N/A`.
- **Employee salary rounding is a bit trivial by design.** All ten salaries are already round multiples of 1,000, so `ROUND`, `CEILING.MATH(..., 100)`, and `FLOOR.MATH(..., 100)` all return the exact same value as the original salary — the formulas are correct, just not visibly doing anything on this particular dataset.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| 🟢 **Microsoft Excel (365/2021+)** | Required for dynamic array functions (`FILTER`, `XLOOKUP`, `XMATCH`) to spill correctly |
| 🧮 **Logical Functions** | `IF`, `IFS`, `AND`, `OR` |
| 🔍 **Lookup Functions** | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`/`MATCH` |
| 💧 **Dynamic Arrays** | `FILTER` |
| 🧵 **Text Functions** | `LEFT`, `FIND`, `UPPER`, `LOWER` |
| 📅 **Date Functions** | `DATEDIF`, `TODAY` |
| 🔢 **Math Functions** | `ROUND`, `CEILING.MATH`, `FLOOR.MATH` |
| 📊 **Aggregate Functions** | `AVERAGE`, `COUNTIF`, `SUMIFS`, `SUM` |
| 🧭 **Reference Functions** | `OFFSET`, `INDIRECT` |

---

## 📈 Results & Insights

- ✅ **3 Independent Sheets** — Student, Sales Data, and Employee Data, none foreign-key linked, each self-contained
- 🎓 **Student Grades** — scores span C through A; every student clears both the >50 and >60 thresholds in both subjects
- 💧 **Top Scorers (80+ average)** — 7 of 10 students qualify by the `FILTER` formula's logic, though the displayed panel currently shows placeholder names (see Notes above)
- 💰 **Total Sales Revenue** — ₹17,11,000 across all 10 transactions
- 🏷️ **Discount Eligibility** — 9 of 10 sales rows qualify for a discount; only the low-price, low-quantity Keyboard sale in May misses both thresholds
- 👔 **Salary Range (Employee Data)** — ₹48,000 (Shruti Das, HR) to ₹98,000 (Simran Kaur, Finance)

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🧮 **Formula-Only Design** | No VBA or Power Query — every result is a live, recalculating cell formula |
| 🔍 **Lookup Function Variety** | Three different lookup approaches (`VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`) solving equivalent problems, useful for comparing trade-offs |
| 💧 **Modern Dynamic Arrays** | `FILTER` demonstrates spill-based formulas rather than legacy array-entered (Ctrl+Shift+Enter) formulas |
| 📊 **Realistic Multi-Sheet Scope** | Three unrelated domains in one file mirrors how real workbooks accumulate multiple reporting tabs |
| 🎓 **Beginner-to-Intermediate Friendly** | Escalates from simple `IF`/`AVERAGE` to multi-criteria array matching and `INDIRECT` referencing |
| 🧪 **Extensible** | Additional sheets (e.g. `Inventory`, `Attendance`) could reuse the same lookup-panel pattern |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Ved Dhameliya

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/isamaliya16)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayush-isamaliya-686533312/)

> *"A well-designed schema is half the application — the formulas just ask it questions."*

**🎓 Role:** Junior Python Developer | Programming Enthusiast \
**📍 Location:** India\
**🛠️ Skills:** Excel · Formula Engineering · Lookup Functions · Dynamic Arrays · Data Cleaning

Video Explanation Link:------

</div>

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: 02 September, 2026*

</div>
