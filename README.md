# CAMeSM Workshop Registration Manager

A lightweight desktop application built to handle workshop registrations for CAMeSM without the usual spreadsheet chaos.

It takes multiple CSV exports (one per workshop), merges them into a single dataset, and helps identify overlaps, attendee profiles, and optimal allocations.

---

## Why this exists

Attendees sign up for workshops through Google Forms, which spits out one CSV per workshop. The problem: people register for multiple workshops, and when a workshop is half-empty while another is overbooked, you need to figure out who to redirect — and email them accordingly.

This app loads all the CSVs, deduplicates by email, builds a profile for each attendee, and tells you exactly who to confirm where.

---

## What it does

- Combines multiple workshop CSV files into a single dataset  
- Groups attendees by email into unified profiles  
- Highlights people registered in multiple workshops  
- Suggests which workshop each person should keep (based on availability)  
- Provides a visual dashboard with real-time statistics  
- Exports clean CSVs for communication and confirmation  

---

## Core features

### 1. CSV ingestion
- One file per workshop  
- Flexible column mapping (email required, others optional)  

### 2. Attendee profiling
- Automatically merges duplicate entries across workshops  
- Tracks number of registrations per person  

### 3. Allocation helper
- Detects conflicts (multi-registrations)  
- Suggests keeping the least-subscribed workshop  
- Helps balance attendance without manual sorting  

### 4. Dashboard
- Registrations per workshop  
- Year distribution  
- Heatmaps and participation breakdowns  

### 5. Export tools
- Full attendee list  
- Allocation suggestions  
- Per-workshop confirmation lists  
- Master registration log  

---

## Tech stack

- Python  
- CustomTkinter (GUI)  
- Pandas  
- Matplotlib  

---

## Installation

```bash
git clone 
cd camasm-workshop-manager
pip install -r requirements.txt
python3 app.py
```

---

## Expected CSV structure

Each file should correspond to one workshop. The filename becomes the workshop name in the app, so name them something sensible (Surgery_Simulation.csv, ECG_Interpretation.csv, etc.).

The app auto-detects columns using keyword matching, so exact header names don't matter much. It looks for:

| Field  | Required | Notes |
|--------|----------|------|
| Email  | Yes      | Unique identifier |
| Name (First)   | No       | Used for display |
| Name (Last)   | No       | Used for display |
| Year   | No       | Used in analytics |

| Field | Recognised keywords |
|---|---|
| Email | `email`, `e-mail`, `mail` |
| First name | `first`, `given`, `forename` |
| Last name | `last`, `surname`, `family` |
| Full name (fallback) | `full`, `name`, `student` |
| Year of study | `year`, `semester`, `study`, `grade` |
 
If detection gets it wrong for a particular file, you can override each column from the dropdown before building. Only the email column is required — everything else is optional.


### Notes

- Column names do not need to match exactly — the app will attempt to detect them automatically.  
- You can override detected columns manually during setup.  
- Empty or inconsistent values (e.g. missing year) are handled and grouped as "Unknown".  
- **Email is used as the primary key**, so duplicates across files are merged automatically.  

---

## Example

A typical CSV exported from a registration form might look like:

| Email              | Name (First) | Name (Last)       | Year       |
|-------------------|--------|---------|------------|
| student1@euc.ac.cy | Maria | Ioannou    | 3rd Year   |
| student2@euc.ac.cy | Andreas | Georgiou | 4th Year   |

Each workshop should have its own CSV file.

---

## Workflow overview

1. Load all workshop CSVs  
2. Map columns (if needed)  
3. Build dataset  
4. Review dashboard and profiles  
5. Resolve multi-registrations via Allocation Helper  
6. Export final lists for communication  

---

## Allocation logic

If someone registers for multiple workshops:

The app ranks their workshops by current unique registrant count and suggests confirming them in the least-subscribed one. It doesn't try to optimise globally — it just makes a greedy per-attendee suggestion based on the current snapshot. Good enough for the actual problem, which is filling workshops that would otherwise run half-empty.
 
The allocation is a suggestion, not a decision. You still send the emails manually.

---

## Context

Built for the Cyprus Annual Medical Student Meeting (CAMeSM) 2026 to streamline workshop coordination at scale.

---

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.

---

*Written by Alexandros Kordatzakis — CAMeSM Workshops Coordinator & IT/Web Developer*
