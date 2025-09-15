# Excel Skills Reference Guide

This guide summarizes the core Excel skills I practiced and how to use them. It is meant as a quick reference for consulting and business analysis work.

---

## 📊 PivotTables

### What it does:
Summarizes and analyzes large datasets quickly.

### How to create:
1. Select your dataset.
2. Go to **Insert → PivotTable**.
3. Place it in a new worksheet.
4. Drag fields into:
   - **Rows** (categories)
   - **Columns** (sub-categories)
   - **Values** (calculations: Sum, Average, Count)
   - **Filters** (filtering)

### Example:
- Average Salary by Department:  
  - Rows → Department  
  - Values → Salary (set to Average)

---

## 🔎 VLOOKUP

### What it does:
Searches for a value in the first column of a table and returns data from another column.

### Syntax:
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

### Example:
```excel
=VLOOKUP("E1007", A2:F50, 4, FALSE)
```
- Looks for Employee_ID `E1007` in column A.
- Returns the value from column 4 (Salary).

---

## 🔎 XLOOKUP

### What it does:
More flexible version of VLOOKUP. Directly specifies lookup and return columns.

### Syntax:
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

### Example:
```excel
=XLOOKUP("E1007", A2:A50, D2:D50, "Not Found")
```
- Looks in column A for `E1007`.
- Returns the Salary from column D.

### Benefits over VLOOKUP:
- No need to count column index.
- Doesn’t break if columns move.

---

## ✅ IF Statements (Conditional Logic)

### Basic IF
```excel
=IF(E2>=4,"High Performer","Standard")
```

### Nested IF
```excel
=IF(E2=5,"Excellent",
   IF(E2=4,"Good",
   IF(E2=3,"Average",
   "Needs Improvement")))
```

### IFS (cleaner)
```excel
=IFS(E2=5,"Excellent", E2=4,"Good", E2=3,"Average", E2<=2,"Needs Improvement")
```

---

## 🎨 Conditional Formatting

### Purpose:
Highlight values or add visual cues.

### Common Uses:
1. **Highlight > 60000 (Salary)**  
   - Select D2:D50 → Conditional Formatting → New Rule.  
   - Condition: `Cell Value > 60000`.  
   - Format: Bold.

2. **Color Scales**  
   - Apply gradient (Green → Red) to Performance_Score.

3. **Icons**  
   - Add arrows/traffic lights for quick insights.

4. **Formula Rule Example**  
   ```excel
   =$F2>=10
   ```
   - Highlights employees with 10+ Years_at_Company.

---

## 📈 Common Functions

- **SUM / AVERAGE / COUNT / MAX / MIN** → aggregate values.
- **SUMIF / COUNTIF** → aggregate with conditions.  
  Example:  
  ```excel
  =COUNTIF(B2:B50,"Sales")
  ```
- **Text Functions** → `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCAT`, `TEXTJOIN`.

---

## 🧹 Data Cleaning Tools

- **Remove Duplicates** → Data tab.  
- **Text-to-Columns** → Split names/regions.  
- **Flash Fill** → Excel predicts patterns (Ctrl+E).  
- **Find & Replace** → Clean up data quickly.

---

## 📊 Charts

- Create PivotCharts (bar, column, pie).  
- Add axis labels and titles for clarity.  
- Use conditional formatting + charts together for client-ready visuals.

---

## ⚡ Efficiency Tips

- **Ctrl + Shift + L** → Toggle filters.  
- **Alt + =** → Auto-sum.  
- **Ctrl + Shift + →/↓** → Select range.  
- **Ctrl + T** → Convert to table.  
- **Freeze Panes** → Keep headers visible.  
- **Data Validation** → Create dropdowns.

---

# ✅ Key Takeaway

For business/consulting interviews, I am comfortable with:
- PivotTables for analysis.
- VLOOKUP/XLOOKUP for lookups.
- IF/IFS for conditional logic.
- Conditional Formatting for insights.
- Core functions and data cleaning.

This makes me capable of analyzing, summarizing, and presenting data effectively in Excel.
