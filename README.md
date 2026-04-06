![Cover Image](https://raw.githubusercontent.com/saifulroni/-automated_project_health_monitor/1e2eb2af41ea081e28fefc62a04c0d975123518f/cover_image.png)

# Automated_project_health_monitor
FMCG Portfolio | Project Automated PMO Dashboard | Python + Excel | 

---

## What This Project Does

This tool automates the weekly project health reporting process for a simulated FMCG product launch (energy drink brand "Lemu Tajin" by Akij Ventures). What a PMO coordinator would spend 3–4 hours doing manually every day is now fully automated.

**Input:** One Excel file (`tasks.xlsx`) updated by the project team  
**Output:** Formatted Excel dashboard + PDF health report, auto-emailed to the PM

---

# Business Context

The tool manages 40 tasks across 3 concurrent workstreams:
- **Supply Chain Readiness** — packaging, production, QA, warehouse, logistics
- **Marketing Campaign Preparation** — creative, media, digital, PR, launch event
- **Distributor & Retail Onboarding** — agreements, training, shelf space, go-live

Every Monday morning the PMO head receives a one-page health summary for each project automatically generated every Monday at 5pm by Windows Task Scheduler.

---

## Features

- **RAG Status Engine** — Green / Amber / Red logic with explicit business rules, recalculated automatically based on date and task progress
- **Critical Path Analysis** — uses NetworkX directed acyclic graph to identify the longest dependency chain; critical path tasks flagged distinctly
- **Gantt Chart Generator** — matplotlib Gantt with RAG color coding, today marker, and critical path highlighting; exported as PNG
- **Excel Dashboard** — formatted openpyxl workbook with portfolio summary, per-project health views, and overdue/blocked task tables
- **PDF Weekly Report** — ReportLab report covering portfolio RAG grid, overdue tasks, at-risk upcoming tasks, and critical path status
- **Automated Email Delivery** — smtplib sends the PDF to the PM every monday via Gmail SMTP; credentials stored in environment variables, never in code
- **Task Scheduler Integration** — Windows Task Scheduler triggers main.py every monday at 9pm with zero manual intervention

---


## Sample Outputs
**Generated Gantt Chart**
![Gantt chart image](https://raw.githubusercontent.com/saifulroni/-automated_project_health_monitor/c28368afeda248d499530f7ba33b8bd9b8343700/outputs/gantt_Distributor_%26_Retail_Onboarding.png)

**Generated Excel Dashboard**
![Dashboard image](https://raw.githubusercontent.com/saifulroni/-automated_project_health_monitor/7561cfc713c5389387d0bac2c110749fc86d8b79/dashboard%20image.png)


**Generated Weekly report**


![weekly report 1](https://raw.githubusercontent.com/saifulroni/-automated_project_health_monitor/4ce0ebd9361bd410b433a1653c46e48ea2867370/weekly%20report%201.png)
![weekly report 2](https://raw.githubusercontent.com/saifulroni/-automated_project_health_monitor/4ce0ebd9361bd410b433a1653c46e48ea2867370/weekly%20report%202.png)


---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.11 | Core language |
| pandas | Data ingestion and transformation |
| NetworkX | Directed acyclic graph for critical path |
| openpyxl | Excel dashboard generation |
| matplotlib | Gantt chart generation |
| ReportLab | PDF report generation |
| smtplib | Automated email delivery |
| Windows Task Scheduler | Weekly automation trigger |


