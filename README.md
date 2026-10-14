# Excel Data Science Salary Dashboard

## Introduction

This Excel dashboard helps users explore salary trends across data-related roles and compare median salaries based on **job title, country, and job schedule type**. It combines a large job-posting dataset with interactive selections, formulas, and charts to turn raw job-market data into practical career insights.

The workbook uses a dataset containing data-related job information from **2023**, including job titles, locations, salary information, schedule types, and skills.

## 🎯 Project Objectives

- Explore median salaries across data and analytics job titles.
- Compare salary levels by country and job schedule type.
- Use formulas to calculate salary metrics from the underlying dataset.
- Present the results in a dashboard that is easier to interpret than raw data alone.

## Dataset Overview
The dataset used for this project contains real-world data science job information from 2023. The workbook contains a structured Excel table named `jobs` on the **Data** worksheet. It contains **32,672 job records** and 16 fields, including:

- **Job information:** `job_title_short`, `job_title`
- **Location:** `job_location`, `search_location`, `job_country`
- **Job arrangement:** `job_schedule_type`, `job_work_from_home`
- **Posting details:** `job_via`, `job_posted_date`
- **Salary fields:** `salary_rate`, `salary_year_avg`, `salary_hour_avg`
- **Additional details:** degree requirements, health insurance, company name, and job skills

The dataset provides the foundation for the salary calculations and dashboard visuals.

## Dashboard Overview

The **Dashboard** sheet is the main user-facing page. It includes dropdown selections for:

- **Job Title**
- **Country**
- **Job Type**

These selections drive the salary analysis so users can explore a particular combination instead of relying on one fixed view.

### 📉 Charts

#### Data Science Job Salaries - Bar Chart

<img src="/0_Resources/Images/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

- 🛠️ **Excel Features:** Utilized bar chart feature (with formatted salary values) and optimized layout for clarity.
- 🎨 **Design Choice:** Horizontal bar chart for visual comparison of median salaries.
- 📉 **Data Organization:** Sorted job titles by descending salary for improved readability.
- 💡 **Insights Gained:** This enables quick identification of salary trends, noting that Senior roles and Engineers are higher-paying than Analyst roles.

#### Country Median Salaries - Map Chart

<img src="/0_Resources/Images/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

- 🛠️ **Excel Features:** Utilized Excel's map chart feature to plot median salaries globally.
- 🎨 **Design Choice:** Color-coded map to visually differentiate salary levels across regions.
- 👁️ **Visual Enhancement:** Improved readability and immediate understanding of geographic salary trends.
- 💡 **Insights Gained:** Enables quick grasp of global salary disparities and highlights high/low salary regions.

### 🧮 Formulas & Excel Techniques

The workbook uses a range of Excel functions and techniques to prepare the analysis.

| Function / technique | Purpose |
|---|---|
| `MEDIAN` | Calculates median annual salary for matching records |
| `IF` | Applies criteria when selecting salary values for calculation |
| `SEARCH` and `ISNUMBER` | Checks whether job schedule text matches the selected type |
| `FILTER` | Excludes schedule-type entries that should not appear in a selection list |
| `UNIQUE` | Builds lists of distinct job titles, countries, and schedule types |
| `SORT` | Sorts salary results to support chart comparisons |
| `XLOOKUP` | Retrieves the salary result associated with a selected job title |
| `COUNTIF` / `COUNTIFS` | Counts records that match specified criteria |
| Structured references | Refers to fields in the `jobs` Excel Table |
| Data Validation | Restricts dashboard selections to available list values |

### Example: Median Salary Calculation

```
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
```

Conceptually, the formula:

1. Checks the job title against the selected title.
2. Checks the country against the selected country.
3. Checks whether the schedule-type field matches the selected type.
4. Excludes zero annual salary values.
5. Calculates the median of the remaining annual salaries.

This allows the displayed result to respond to the dashboard selections rather than using a single overall salary figure.

### ❎ Data Validation

#### 🔍 Filtered List

- 🔒 **Enhanced Data Validation:** Implementing the filtered list as a data validation rule under the `Job Title`, `Country`, and `Type` option in the Data tab ensures:
    - 🎯 User input is restricted to predefined, validated schedule types
    - 🚫 Incorrect or inconsistent entries are prevented
    - 👥 Overall usability of the dashboard is enhanced

<img src="/0_Resources/Images/1_Salary_Dashboard_Data_Validation.gif" width="425" height="400" alt="Salary Dashboard Data Validation">

## 🔄 How the Dashboard Works

The workflow is designed to be straightforward:

```text
1. Select a job title, country, and job type
                  ↓
2. Data validation keeps selections consistent
                  ↓
3. Formula-driven calculation sheets filter the jobs dataset
                  ↓
4. Median salary metrics are calculated from matching records
                  ↓
5. Supporting tables sort and prepare the results
                  ↓
6. Dashboard charts present the salary comparisons
```

The user mainly interacts with the dropdowns on the **Dashboard** sheet. The supporting worksheets contain the lists, formulas, and intermediate calculations used to generate the results.

## 🗂️ Workbook Structure

The workbook separates raw data, user selections, calculations, and reporting across multiple worksheets.

| Worksheet | Purpose |
|---|---|
| `Data` | Main job-postings dataset stored in the `jobs` Excel Table |
| `Dashboard` | Main interface with selection dropdowns and charts |
| `Validation` | Lists and calculation ranges used to support dropdowns and filtering |
| `Salary` | Median salary calculations and sorted job-title salary results |
| `Country` | Country-related salary data and supporting results |
| `Type` | Supporting calculations for job schedule-type comparisons |
| `Platform` | Supporting job posting platform counts and calculations |

## 🛠 Tools & Skills Demonstrated

### Tools

- **Microsoft Excel**
- Excel Tables
- Excel Formulas and Functions
- Data Validation
- Dynamic array formulas
- Charts and dashboard design

### Excel & Analytics Skills

- Data organization and structured references
- Multi-criteria filtering
- Median salary analysis
- Dynamic list generation
- Salary comparisons by job title, country, and schedule type
- Formula-driven reporting
- Chart preparation and dashboard visualization
- Separating raw data, calculations, and presentation

## How to Use

1. Open `Final_Dashboard.xlsx` in Microsoft Excel.
2. Go to the **Dashboard** worksheet.
3. Use the dropdowns to select a **Job Title**, **Country**, and **Job Type**.
4. Review the salary result and charts for the selected criteria.
5. Change one or more selections to compare a different role, location, or schedule type.

The workbook uses modern Excel functions, including dynamic array functions. For best compatibility, open it in a recent version of Microsoft Excel that supports these functions.

## 💡 Design Principles

This project is built around four principles:

**1. Relevant comparisons**  
Salary results are calculated for selected criteria so users can compare roles and employment contexts.

**2. Formula-driven analysis**  
Excel formulas automate filtering, median calculations, sorting, and lookups.

**3. Consistent inputs**  
Dropdown-based data validation helps users choose values from predefined lists.

**4. Separation of logic and presentation**  
The dataset and supporting calculation worksheets power the dashboard, while the Dashboard sheet provides the main interface for exploration.

## Potential Future Improvements

The dashboard could be extended with:

- 📅 Salary trends over time using posting dates
- 🧰 Skill-based salary comparisons
- 🔎 Slicers or other interactive filtering options
- 📊 Additional KPIs, such as the number of matching salary records

## 🧠 What I Learned

This project provided practical experience in building an Excel-based salary analysis dashboard.

Key takeaways include:

- Organizing a large dataset in an Excel Table
- Using formulas to calculate median salaries across multiple criteria
- Creating dynamic lists for dashboard selections
- Applying data validation to guide user input
- Preparing sorted results for charting
- Connecting calculations to a user-facing dashboard
- Turning raw job-posting data into a more accessible salary comparison tool

## 👤 Project Purpose

This project demonstrates **Excel data analysis, formula automation, data validation, dynamic filtering, salary analysis, and dashboard development**. It shows how Excel can be used to explore job-market data and present salary comparisons in a practical, interactive format.

## ⭐️ If You Find This Project Useful

Feel free to explore the workbook, adapt the dashboard to your own salary research, or use its structure as a starting point for building other Excel-based analytics projects.
