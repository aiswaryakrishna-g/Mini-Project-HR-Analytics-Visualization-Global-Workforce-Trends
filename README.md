[README.md](https://github.com/user-attachments/files/32213383/README.md)
# Mini-Project-HR-Analytics-Visualization-Global-Workforce-Trends
Mini-Project // HR Analytics &amp; Visualization-Global Workforce Trends //Using POWERBI
# Global Workforce Trends — HR Analytics & Visualization (2014–2025)

An end-to-end HR analytics project that simulates a global company's workforce and turns raw, messy data into a polished, interactive **Power BI** dashboard. The project covers the full pipeline: data modeling, cleaning in Excel/Power Query, DAX measure design, and a 5-page executive-ready report with drillthrough and insight-driven recommendations.

---

## 📌 Objective

This dataset simulates a global company's workforce with thousands of fictional employees across multiple countries and 25+ departments. It includes demographic details, departments, locations, promotions, annual salaries, and hierarchical manager–employee relationships — enabling both cross-sectional and time-series analysis.

**Example use cases:**
- HR analytics and workforce demographics
- Salary modeling and pay equity studies
- Career growth and promotion trend analysis
- Organizational network visualization

---

## 📂 Dataset Overview

| File | Description |
|---|---|
| `employees.xlsx` | Core demographic and organizational attributes |
| `departments.xlsx` | Department names and categories |
| `locations.xlsx` | Global office locations with cost-of-living indices |
| `promotions.xlsx` | Career progression events over time |
| `salaries_annual.xlsx` | Yearly compensation history (2014–2025) |
| `org_edges.xlsx` | Manager–employee relationships for org chart visualization |

---

## 🧹 Data Pre-Processing (Excel & Power Query)

### Employees Table
- **Deduplication:** Checked for duplicate employee IDs; used `Data → Remove Duplicates` to ensure unique records across all tables.
- **Standardization:**
  - Capitalized text fields via Power Query (`Format → Capitalize Each Word`)
  - Applied a consistent date format (`dd-mm-yyyy`) via `Power Query → Change Type → Date`
  - Removed extra spaces (`TRIM`) and non-printable characters (`CLEAN`) via Power Query
  - Converted salary columns to Number format for calculations
  - Found 66 blank cells using `=COUNTBLANK()` (missing `Manager Id` with no match in the `org_edges` table) and replaced them with `"Unknown"` via Find & Replace
  - Pulled the `Continent` column from the Locations table using `XLOOKUP`:
    ```excel
    =XLOOKUP([@City], locations[city], locations[Continent])
    ```
  - Verified no null values remained (`Ctrl+F` audit)
  - Centered text alignment across the table

- **Calculated Fields:**
  - **Tenure:**
    ```excel
    =DATEDIF([@[hire_date]], TODAY(), "y")
    ```
  - **Cost Index** (brought from Locations table via `VLOOKUP`):
    ```excel
    =VLOOKUP([@[location_id]], locations[#All], 4, )
    ```
  - **Department Name / Category** (brought from Departments table via `VLOOKUP`):
    ```excel
    =VLOOKUP([@[department_id]], departments, 2, FALSE)
    =VLOOKUP([@[department_id]], departments, 3, FALSE)
    ```
  - **Cost Adjusted Salary** — normalizes pay across countries by removing local cost-of-living bias:
    ```excel
    =[@[initial_salary_usd]] / [@[cost_index]]
    ```

### Departments, Locations & Org Edges Tables
- Checked for duplicates (none found in any of these tables)
- Removed extra spaces (`TRIM`) and non-printable characters (`CLEAN`) via Power Query
- Verified no null values; centered text alignment
- Added a `Continent` column to the Locations table

### Promotions Table
- Capitalized text in the `Level` / `To Level` columns via Power Query
- Verified no null values; centered text alignment

### Salaries Annual Table
- Converted the salary column to Decimal format
- Verified no null values; centered text alignment
- **Calculated Field — Annual Salary Growth %:**
  ```excel
  =IF(A8=A7, (D8-D7)/D7, 1)
  ```
  *(Assumes the first year is the "base" year and sets growth to 100%.)*

### Pivot Tables & Charts
Built pivot tables to summarize:
- Employee count by department
- Average salary by department
- Employee count by location
- Gender distribution
- Experience level vs. salary

### Filtering & Sorting
- Removed irrelevant test entries and incomplete records prior to loading into Power BI.

---

## 📊 Power BI Visualization Dashboard

### Report Structure
A **5-page report** connected by a page-navigator button set, with a **Reset** button available on every page:

```
Overview → Compensation → Career & Org → Insights & Recommendations → Drill Through
```

### Data Model
- **Measure Table** — a dedicated disconnected table holding all central DAX measures
- **Lowest Satisfied Emp Grouped by Performance** — a calculated/grouped table (via `GROUPBY`) used to analyze the lowest-satisfaction employees segmented by performance tier

**Calculated column — Promotion Count:**
```dax
Promotion count =
CALCULATE(
    COUNTROWS(promotions),
    FILTER(promotions, promotions[employee_id] = EARLIER(promotions[employee_id]))
)
```

**Grouped table — Lowest Satisfied Emp Grouped by Performance:**
```dax
Lowest Satisfied Emp Grouped by Performance =
GROUPBY(
    employees,
    employees[performance_score],
    "Min of satisfaction_score", MINX(CURRENTGROUP(), employees[satisfaction_score]),
    "Avg Salary", AVERAGEX(CURRENTGROUP(), employees[Cost Adjusted Salary])
)
```

### Visualizations by Page

**1. Overview**
- 5 KPI cards: Total Employees, Average Salary, Average Tenure, Average Performance, Average Satisfaction
- Donut chart: headcount by department category
- Pie chart: gender distribution
- Line chart: hires per year (using the Year level of a date hierarchy on `hire_date`)
- Map: employee count by country
- Gauge: average performance
- Pivot table: top 10 employees by performance
- 5 slicers + page navigator

**2. Compensation**
- 3 KPI cards: Total Payroll, Median Cost-Adjusted Salary, Salary Growth Trend
- Treemap: average salary split by male/female/non-binary
- Line chart: top-paid department over time
- Combo chart: salary by seniority
- Bar chart: salary by education
- Waterfall chart: salary by tenure
- 5 slicers

**3. Career & Org**
- 3 KPI cards: Total Promotions, Managers, Employees per Manager
- Clustered column chart: promotions by year × level (Lead / Manager / Mid / Senior)
- Pie chart: headcount by work mode
- Pivot table: employees × promotion count × performance
- 5 slicers

**4. Insights & Recommendations**
- 3 KPI cards: Lowest Satisfaction Score, Managers, Lowest Cost Index
- Scatter chart: performance vs. cost index by continent
- Pivot table sourced from the disconnected grouped table

**5. Drill Through**
- Detail table: department category, department name, seniority, work mode, city, employee count
- Accessible via right-click drillthrough from any other page

---

## 🧮 Key DAX Measures

```dax
Average Tenure = FORMAT(AVERAGE(employees[Tenure]), "0.00") & " years"

avg performance = AVERAGE(employees[performance_score])

Avg Salary Female =
CALCULATE(AVERAGE(employees[initial_salary_usd]), Employees[Gender] = "Female")

Avg Salary Male =
CALCULATE(AVERAGE(employees[initial_salary_usd]), Employees[Gender] = "Male")

Avg Salary NonBinary =
CALCULATE(AVERAGE(employees[initial_salary_usd]), Employees[Gender] = "Non-binary")

count of satisfied employee = COUNT(employees[name])

Employee count = COUNTROWS(employees)

employees per manager = COUNTROWS(RELATEDTABLE(org_edges))

Lowest Cost Index = MIN(employees[cost_index])

Lowest Satisfaction by Manager =
CALCULATE(
    MIN(Employees[satisfaction_score]),
    ALLEXCEPT(Employees, Employees[manager_id])
)

Median of Cost adjusted salary = MEDIAN(employees[Cost Adjusted Salary])

salary growth percentage =
DIVIDE(
    (SUM(employees[initial_salary_usd]) - CALCULATE(SUM(employees[initial_salary_usd]), DATEADD(employees[hire_date], -1, YEAR))),
    CALCULATE(SUM(employees[initial_salary_usd]), DATEADD(employees[hire_date], -1, YEAR))
)

Top Paid Dep =
MAXX(VALUES(employees[department_category]), CALCULATE(AVERAGE(employees[initial_salary_usd])))

Total Payroll = SUM(employees[initial_salary_usd])
```

---

## 💡 Insights

1. **Africa is the efficiency leader** — near-top performance at the second-lowest cost index, the clearest "do more of this" signal in the data.
2. **Europe is the outlier to investigate** — paying a premium (cost index > 1.1) without a matching performance return; of all six regions, it has the worst cost-to-output ratio.
3. **Oceania's premium looks justified** — the most expensive region, but also posts the single highest performance score, making it a defensible cost rather than a wasteful one.
4. **South America is "cheap but unproven,"** not yet "cheap and strong" like Africa — worth separating these two narratives, since low cost alone isn't the win; low cost *with* performance is.

## ✅ Recommendations

| Priority | Recommendation |
|---|---|
| 🔴 High | **Investigate Europe's performance gap** — highest cost tier, weakest performance. Audit hiring bar, management quality, or role mix before renewing headcount plans. |
| 🔴 High | **Prioritize incremental headcount growth in Africa** — best cost-to-performance ratio on the chart; scaling here has the most efficient payoff. |
| 🟡 Medium | **Run a root-cause review on the lowest-satisfaction segment** — a floor score of 5.00 needs an average/distribution view alongside it to determine if it's one outlier or a systemic dip. |
| 🟡 Medium | **Pilot targeted development support in South America** — the cost advantage is real, but performance needs to catch up before treating it like Africa. |

---

## 🛠️ Tools Used

- **Microsoft Excel** — data cleaning, Power Query transformations, pivot tables, calculated fields (XLOOKUP, VLOOKUP, DATEDIF)
- **Power BI** — data modeling, DAX measures, interactive report design (KPI cards, maps, treemaps, waterfall/combo charts, drillthrough, bookmarks/navigation)

---


```

---

## 📬 Contact

Feel free to open an issue or reach out if you have questions about the data model, DAX logic, or dashboard design choices.
