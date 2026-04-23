# CAMeSM Workshop Registration Manager

A lightweight desktop application built to handle workshop registrations for CAMeSM without the usual spreadsheet chaos.

It takes multiple CSV exports (one per workshop), merges them into a single dataset, and helps identify overlaps, attendee profiles, and optimal allocations.

---

## Why this exists

Managing workshop registrations manually becomes messy very quickly:
- duplicate sign-ups  
- inconsistent CSV formats  
- uneven workshop distribution  

This tool was written to solve that in a way that is fast, visual, and practical during event preparation.

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
pip install -r requirements.txt
python app.py
```

---

## Expected CSV structure

Flexible — only one requirement:

| Field  | Required | Notes |
|--------|----------|------|
| Email  | Yes      | Unique identifier |
| Name (First)   | No       | Used for display |
| Name (Last)   | No       | Used for display |
| Year   | No       | Used in analytics |

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

- Workshops are ranked by number of attendees  
- The attendee is assigned to the least-subscribed one  
- Other registrations remain visible for reference  

This helps balance attendance without forcing manual decisions.

---

## Context

Built for the Cyprus Annual Medical Student Meeting (CAMeSM) to streamline workshop coordination at scale.

---

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.
