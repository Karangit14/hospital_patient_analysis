# Hospital Emergency Room Dashboard -- Excel

## 📊 Project Overview

The **Hospital Emergency Room Dashboard** is an interactive Microsoft
Excel dashboard designed to analyze hospital emergency-room patient data
and present key operational insights in a simple visual format.

The dashboard uses **Excel PivotTables, DAX/Power Pivot measures, linked
PivotTable outputs, charts, and interactive filters** to create a
monthly patient-analysis report.

The main purpose of the project is to help hospital management
understand patient volume, waiting time, satisfaction, admission status,
gender distribution, age distribution, and departmental referrals.

------------------------------------------------------------------------

## 🎯 Project Objectives

-   Analyze the number of emergency-room patients.
-   Monitor average patient waiting time.
-   Track average patient satisfaction score.
-   Compare admitted and non-admitted patients.
-   Analyze patient attendance status.
-   Understand gender distribution.
-   Analyze patients by age group.
-   Identify departmental referral patterns.
-   Provide an interactive monthly reporting dashboard.

------------------------------------------------------------------------

## 🛠️ Tools & Technologies

  -----------------------------------------------------------------------
  Tool                                Purpose
  ----------------------------------- -----------------------------------
  **Microsoft Excel**                 Data preparation and dashboard
                                      development

  **Power Query / Excel Data Tools**  Data cleaning and transformation,
                                      where applicable

  **Power Pivot / DAX**               Measures and analytical
                                      calculations

  **PivotTables**                     Aggregation and analysis

  **PivotCharts / Excel Charts**      Data visualization

  **Excel Slicers / Filters**         Interactive filtering
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 🔄 Dashboard Workflow

``` text
Raw Hospital Patient Data
          ↓
Data Cleaning & Preparation
          ↓
Excel Data Model / Power Pivot
          ↓
DAX Measures
          ↓
PivotTables
          ↓
Linked PivotTable / Chart Outputs
          ↓
Interactive Excel Dashboard
```

------------------------------------------------------------------------

## 📌 Dashboard Features

### 1. Monthly Filter

The left-side month selector allows the user to analyze patient
information for individual months from **January to December**.

A year selector is also provided for switching between available
reporting years.

### 2. Total Number of Patients

Displays the total number of patients for the selected month/year.

Example shown in the dashboard:

**488 patients**

### 3. Average Waiting Time

Shows the average time patients waited in the emergency room.

Example shown:

**35.20 minutes**

### 4. Average Satisfaction Score

Displays the average patient satisfaction score.

Example shown:

**4.79**

### 5. Attendance Status

A doughnut/pie-style chart compares:

-   **Delay**
-   **Ontime**

This helps identify the proportion of patients who received delayed
versus on-time attention.

### 6. Gender Analysis

The gender chart compares:

-   Female
-   Male

This provides a quick view of the gender composition of emergency-room
patients.

### 7. Admission Status

The dashboard compares:

-   **Admitted**
-   **Not admitted**

The displayed example shows:

-   Not admitted: **222**
-   Admitted: **266**

### 8. Age Group Analysis

A column chart displays patient counts across age groups:

-   0--09
-   10--19
-   20--29
-   30--39
-   40--49
-   50--59
-   60--69
-   70--79

This helps identify which age groups contribute most to emergency-room
visits.

### 9. Departmental Referrals

A horizontal bar chart shows the number of patients associated with
different departments/referral categories, including:

-   Renal
-   Neurology
-   Cardiology
-   Gastroenterology
-   Physiotherapy
-   Orthopedics
-   General Practice
-   None

This can help management understand referral demand across departments.

------------------------------------------------------------------------

## 🧮 DAX / Data Model

The dashboard is supported by a Power Pivot/DAX-based analytical layer.

Typical measures used for this type of dashboard include:

### Total Patients

``` dax
Total Patients = COUNTROWS(PatientData)
```

### Average Waiting Time

``` dax
Average Waiting Time = AVERAGE(PatientData[Waiting Time])
```

### Average Satisfaction Score

``` dax
Average Satisfaction = AVERAGE(PatientData[Satisfaction Score])
```

### Admitted Patients

``` dax
Admitted Patients =
CALCULATE(
    COUNTROWS(PatientData),
    PatientData[Admission Status] = "Admitted"
)
```

### Non-Admitted Patients

``` dax
Non Admitted Patients =
CALCULATE(
    COUNTROWS(PatientData),
    PatientData[Admission Status] = "Not Admitted"
)
```

> **Note:** The exact DAX formulas should be adjusted to match the
> actual table and column names in the Excel workbook.

------------------------------------------------------------------------

## 📈 Key KPIs

The dashboard focuses on the following healthcare KPIs:

  KPI                      Purpose
  ------------------------ ---------------------------------------------
  Total Patients           Measures emergency-room patient volume
  Average Waiting Time     Monitors service efficiency
  Average Satisfaction     Measures patient experience
  Admission Status         Compares admitted and non-admitted patients
  Attendance Status        Tracks delayed vs on-time attention
  Gender Distribution      Understands patient demographics
  Age Distribution         Identifies high-volume age groups
  Departmental Referrals   Understands referral patterns

------------------------------------------------------------------------

## 🎨 Dashboard Design

The dashboard follows a management-reporting layout with:

-   KPI cards at the top
-   Interactive month/year filters
-   Pie/doughnut charts
-   Column chart for age groups
-   Horizontal bar chart for departments
-   Linked PivotTable analysis
-   Consistent dashboard formatting
-   Clear section titles and labels

The design is intended to make important hospital metrics understandable
at a glance.

------------------------------------------------------------------------

## 📂 Suggested Project Structure

``` text
Hospital-Emergency-Room-Dashboard/
│
├── Hospital_Emergency_Room_Dashboard.xlsx
├── README.md
├── dataset/
│   └── hospital_patient_data.xlsx
│
└── screenshots/
    └── hospital_dashboard.png
```

------------------------------------------------------------------------

## 🚀 How to Use the Dashboard

1.  Open the Excel workbook.
2.  Enable editing/content if Excel asks for permission.
3.  Open the dashboard sheet.
4.  Select the required **year**.
5.  Select the required **month**.
6.  Review the KPI cards.
7.  Analyze admission, attendance, gender, age, and departmental charts.
8.  Use PivotTables for detailed analysis when required.
9.  Refresh the data model/PivotTables after updating the source data.

### Refreshing the Dashboard

If the source data changes:

1.  Update or replace the source dataset.
2.  Refresh the Power Query/data connection if used.
3.  Refresh the Power Pivot/Data Model.
4.  Refresh the PivotTables.
5.  Verify that the dashboard charts reflect the updated values.

------------------------------------------------------------------------

## 💡 Business Insights This Dashboard Can Provide

Hospital management can use the dashboard to answer questions such as:

-   How many patients visited the emergency room?
-   What is the average waiting time?
-   Is patient satisfaction improving?
-   What percentage of patients were admitted?
-   Are patients receiving attention on time?
-   Which gender represents a larger share of patients?
-   Which age group has the highest patient volume?
-   Which departments receive the most referrals?
-   How does patient activity change from month to month?

------------------------------------------------------------------------

## 📌 Business Value

This dashboard converts raw hospital patient records into an interactive
management-reporting tool.

It can support:

-   Emergency-room performance monitoring
-   Patient-flow analysis
-   Service-quality monitoring
-   Resource planning
-   Departmental workload analysis
-   Management decision-making

------------------------------------------------------------------------

## 🔮 Possible Future Improvements

The dashboard can be enhanced by adding:

-   Year-over-year patient comparison
-   Monthly trend analysis
-   Peak-hour analysis
-   Average waiting time by department
-   Satisfaction score by department
-   Admission rate by age group
-   Conditional formatting for critical KPIs
-   Automated data refresh
-   Additional drill-down reports
-   Forecasting of emergency-room patient volume

------------------------------------------------------------------------

## 👨‍💻 Project Type

**Data Analytics / Business Intelligence Project**

**Platform:** Microsoft Excel\
**Techniques:** DAX, Power Pivot, PivotTables, PivotCharts, Data
Visualization\
**Domain:** Healthcare / Hospital Emergency Room Analytics

------------------------------------------------------------------------

## 📄 Dashboard Preview

The dashboard provides an interactive monthly hospital emergency-room
report containing KPI cards, patient demographics, admission analysis,
age-group analysis, attendance status, and departmental referral
analysis.

------------------------------------------------------------------------

## ⭐ Skills Demonstrated

-   Excel Data Analysis
-   Data Cleaning
-   Power Pivot
-   DAX
-   PivotTables
-   PivotCharts
-   KPI Development
-   Dashboard Design
-   Healthcare Analytics
-   Business Intelligence
-   Data Visualization
