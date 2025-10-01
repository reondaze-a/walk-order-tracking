# 📦 Walk Order Tracking Automation

An internal automation system I built to streamline **Walk Order (WO)** processing in a warehouse environment.  
The system leverages **Power Automate**, **Excel Office Scripts**, and **SharePoint** to parse standardized emails into structured logs, validate data, and update statuses automatically — reducing manual input and errors.

---

## 🚀 Project Overview
Before this automation, walk orders were logged manually from inconsistent email replies.  
This was error-prone and time-consuming, with no central “source of truth.”  

My solution automates the workflow:
1. **Email Trigger (Power Automate)** – listens for incoming WO-related emails.  
2. **Parser** – extracts carton numbers, SKUs, barcodes, units, locations, and quantities.  
3. **Excel Office Scripts** – logs WO#, enforces data validation, updates statuses.  
4. **SharePoint Integration** – maintains centralized logs accessible across teams.  
5. **Backups & Monitoring** – daily snapshots for historical tracking.

---

## 🔑 Key Features
- 📩 **Automated Email Parsing** – extracts WO details directly from the email body.  
- 📝 **Data Validation** – ensures WO_ID uniqueness and consistent formats.  
- 📊 **Excel + SharePoint Logs** – single source of truth for WO statuses.  
- 🔄 **Status Updates** – replies from pick teams update WO logs automatically.  
- 🕒 **Daily Backups** – midnight snapshots of WO table for audit.  

---

## 🛠 Tech Stack
- **Power Automate** – flows for triggers, parsing, and branching logic.  
- **Excel Office Scripts (TypeScript)** – custom logging, table updates, return objects.  
- **SharePoint** – centralized repository for WO logs and versioning.  
- **Outlook** – standardized WO email formatting.  

---

## 📐 System Workflow
```mermaid
flowchart TD
    A[Email Received] --> B[Power Automate Trigger]
    B --> C[Parse Subject & Body]
    C --> D[Excel Office Script]
    D --> E[SharePoint Log Update]
    E --> F[Backup & Dashboard Refresh]

