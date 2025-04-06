# 🏅 Fairness Analysis of the Olympics (1976–2008)

This project explores the fairness of the Summer Olympic Games from 1976 to 2008, focusing on the impact of **gender inclusion** and **economic disparity** on medal distribution. Leveraging the power of **Excel**, **VBA**, and **macros**, this analysis showcases advanced data manipulation and visualization techniques to uncover key insights.

---

## 📊 Project Highlights

### ✔️ Advanced Excel Skills Demonstrated:
- **Pivot Tables** for dynamic data summarization.
- **Interactive Slicers** linked to multiple charts for quick filtering by year, country, or medal type.
- **Card-style KPI visuals** for displaying key stats like total medals per region or year.
- **Conditional formatting** for visual emphasis on outliers and patterns.
- **Text box linking** to live formula results for dynamic dashboard summaries.

### 🤖 VBA + Macros Automation:
- **Shape-triggered Navigation:**  
  Custom VBA code allows users to click on shapes to jump to different year-specific sheets for faster navigation.
- **Sheet Duplication Macro:**  
  Automatically creates new analysis sheets for any year or scenario using VBA.
- **Custom Button Styling and Slicer Font Scaling:**  
  Macros used to modify UI elements for better dashboard usability and accessibility.
- **Data-driven Slicer Creation:**  
  Code to connect slicers across charts and pivot tables, ensuring unified filtering across the worksheet.
- **Searchable Slicer Simulation:**  
  Simulated a slicer search experience through form controls and search logic using VBA.

---

## 🔍 Analytical Themes

### 📌 Gender Analysis
- Visual trend lines of **female participation growth** over the decades.
- Highlighting persistent **gender gaps in certain sports** (e.g., canoeing, football, sailing).

### 💰 Economic Disparity
- Classification of countries by economic tier (First, Second, Third World).
- Comparative analysis of **medals won vs. GDP** and investment in sports infrastructure.
- Data-driven recommendations for **global equity in athlete development**.

---

## 🧰 Tools & Technologies
| Tool      | Use |
|-----------|-----|
| **Excel** | Data modeling, charts, slicers, dashboards |
| **VBA**   | Sheet automation, macro logic, shape actions |
| **Macros**| UI automation, slicer control, font adjustments |
| **PowerPoint** | Presentation of findings and insights |

---

## 📎 Sample VBA Snippet
```vba
Sub GoToYearSheet()
    Dim year As String
    year = "2000"
    Sheets(year).Activate
End Sub
