# <img src="https://github.com/user-attachments/assets/1cdf5295-d760-42fb-9953-a14decfc6d0b" width="40"/> Data Cleaning & Analysis Using MS Excel Functions
 Hands-On No-Code Data Cleaning, Standardization & Statistical Analysis Using Microsoft Excel on a Real-World Internship Eligibility Dataset

##  <img src="https://github.com/user-attachments/assets/d91c2841-14ca-4283-a7fc-a93fc1e996af" height="22px" style="vertical-align:text-bottom;"> Project Overview

This repository demonstrates end-to-end data cleaning and analysis using MS Excel on a structured student internship eligibility dataset.
The focus areas include:
- Cleaning inconsistent department names
- Standardizing categorical values
- Handling percentages and scores
- Computing weighted scores
- Determining eligibility status
- Generating statistical insights

This project reflects industry-style Excel usage commonly required in operations, analytics, and product roles.

## <img src="https://github.com/user-attachments/assets/f3dcee8e-e008-457a-97fb-d3848b425713" width="25px" style="vertical-align:text-bottom;"> Dataset File
📎 Excel Workbook (Cleaned & Processed)
[_Rutvik_Barbhai_-_225805222_-PlivoProdOps .xlsx](https://github.com/user-attachments/files/24402747/_Rutvik_Barbhai_-_225805222_-PlivoProdOps.xlsx)


## 📊 Student Internship Eligibility Dataset (Sample Preview)

| Student ID | Name | Department | Std. Dept | T1 | T2 | T3 | Attendance | Application Date | Internship Pref | Assignment % | Weighted Score | Status |
|-----------|------|------------|-----------|----|----|----|------------|------------------|------------------|--------------|----------------|--------|
| S001 | Aadhya Jain | CSE | Computer Science | 75 | 59 | 65 | 77% | 2025-08-24 | Company C | 90% | 64.40 | Eligible |
| S002 | Myra Chatterjee | Comp Sci | Computer Science | 66 | 75 | 70 | 86% | 2025-06-08 | Company B | 82% | 71.15 | Eligible |
| S003 | Diya Gupta | CS | Computer Science | 98 | 67 | 82 | 91% | 2025-08-24 | Company D | 79% | 79.15 | Eligible |
| S004 | Sai Agarwal | Computer Science | Computer Science | 96 | 69 | 76 | 99% | 2025-07-27 | Company B | 84% | 76.55 | Eligible |
| S005 | Anika Mukherjee | ECE | Electronics | 78 | 91 | 96 | 81% | 2025-09-15 | Company A | 73% | 91.55 | Eligible |


## <img src="https://github.com/user-attachments/assets/308b07fa-0512-499a-aae7-be67cb7594c3" width="40"/>  Data Cleaning & Standardization 

## <img src="https://github.com/user-attachments/assets/3a0605b5-6c97-40ff-8bb5-af3306c81beb" width="25"/> Application_Raw Spreadsheet

### 🔹 Department Standardization

To handle inconsistent department naming (CSE, CS, Comp Sci, etc.), a lookup-driven standardization was applied.

**Formula used:**
```excel
=VLOOKUP(C2, Dept_Lookup!$A$2:$B$17, 2, FALSE)
```
- Ensures consistent department naming
- Prevents analytical duplication
- Enables accurate aggregation & reporting

### <img src="https://github.com/user-attachments/assets/78993039-3f51-4bbf-b196-9336377732e4" width="24" style="position: relative; top: 3px;"/> Eligibility Status Logic

Eligibility was calculated using attendance + assignment thresholds, dynamically driven by company-specific criteria.
Formula used:
```excel
=IF(
  AND(
    H3 >= (VLOOKUP(J3, Interview_Slots!$A$2:$D$5, 4, FALSE) / 100),
    L3 >= VLOOKUP(J3, Interview_Slots!$A$2:$D$5, 3, FALSE)
  ),
  "Eligible",
  "Not Eligible"
)
```
- Company-wise rules
- No hardcoded values
- Scalable to additional companies

## <img src="https://github.com/user-attachments/assets/3a0605b5-6c97-40ff-8bb5-af3306c81beb" width="25"/> Allocation Spreadsheet

### 🔹 Rank Calculation (Within Department)
Students were ranked department-wise based on weighted score.
Formula used:
```excel
=IFERROR(
  RANK(E3, FILTER($E$2:$E$1000, $D$2:$D$1000 = D3), 0),
  98
)
```
- Fair intra-department ranking
- Dynamic filtering
- Error-handled for edge cases

 ### 🔹Allocation Status

Final allocation is determined based on rank vs available interview slots.
Formula used:
```excel
=IF(
  G5 <= VLOOKUP(D5, Interview_Slots!$A$2:$B$5, 2, FALSE),
  "Allocated",
  "Not Allocated"
)
```
- Slot-aware allocation
- Realistic hiring workflow simulation

##  <img src="https://github.com/user-attachments/assets/6467a42f-afae-4dc7-8ede-46f12a087c6f" width="20" /> Data Visualization & Insight Dashboard

- This section highlights the visual analytics layer of the project, designed to transform cleaned operational data into actionable insights for decision-making teams.
---

### 📊 Chart 1: Best Performing Departments by Average Score
<img src="https://github.com/user-attachments/assets/d91c2841-14ca-4283-a7fc-a93fc1e996af" height="22px" style="vertical-align:text-bottom;"> **Data Flow:** 
```excel
application_raw → Pivot Table → Bar Chart
```
<img width="378" height="211" alt="image" src="https://github.com/user-attachments/assets/c09e815f-3b9b-4859-8338-226aae7cc405" />

- Compares **average scores across standardized departments**
- Highlights **top-performing academic departments**
- Includes a **Grand Total benchmark** for overall comparison


### <img src="https://github.com/user-attachments/assets/6e4a4529-67e1-4729-9ee7-005c5ae3f073" width="24"/> Chart 2: Company Preference Distribution
<img src="https://github.com/user-attachments/assets/d91c2841-14ca-4283-a7fc-a93fc1e996af" height="22px" style="vertical-align:text-bottom;"> **Data Flow:**  
```excel
application_raw → Pivot Table → Pie Chart
```
- Visualizes **student internship preferences across companies**
- Identifies the **most preferred recruiters**
- Helps **operations teams plan interview slots & capacity**
<img width="383" height="236" alt="image" src="https://github.com/user-attachments/assets/62cfbcf1-b372-4c5f-8e9f-76e8df24d688" />

### 📊 Key Insights from the Dashboard

- **Computer Science** consistently shows strong academic performance  
- **Company B** receives the highest preference share among applicants  
- Allocation logic ensures **fair, rank-based selection**
- Dataset and dashboard are **fully scalable** for larger cohorts


### <img src="https://github.com/user-attachments/assets/612137fd-b2de-411c-acd7-f94c4811e9f2" height="25px" style="vertical-align:text-bottom;"> Skills Demonstrated

- Advanced **MS Excel formulas**
- **Lookup-driven standardization**
- Conditional logic & ranking systems
- Operations-style **allocation frameworks**
- Dashboard creation & data storytelling
- Data cleaning & analytics best practices



