---
title: P6 Schedule Comparison
emoji: 📊
colorFrom: indigo
colorTo: blue
sdk: docker
pinned: false
app_port: 7860
---

# P6 Schedule Comparison

Upload two Primavera P6 XER files (baseline + updated) and get an instant 10-tab Project Controls dashboard.

## Tabs

| Tab | What you get |
|-----|-------------|
| Overview | Summary stats, AI health assessment, top delayed activities |
| Schedule | Full activity variance table with search & filter |
| KPIs | Float consumption, weighted % complete, SPI, schedule delay |
| Lookahead | Overdue / 2-week / 4-week / 6-week activity windows |
| Milestones | Milestone tracker with late / at-risk / on-track badges |
| Procurement | Auto-detected procurement WBS items |
| Earned Value | EV metrics (BAC, SPI, CPI, EAC) + S-curves |
| WBS | WBS hierarchy summary |
| Logic | Relationship (predecessor) changes |
| AI Chat | Ask questions about the schedule via streaming AI |

## Local setup

```bash
cd schedule_comparison
pip install -r requirements.txt
cp .env.example .env   # fill in OPENROUTER_API_KEY
source .env
python app.py          # http://localhost:5000
```

## Environment variables (Space secrets)

| Variable | Required | Description |
|----------|----------|-------------|
| `OPENROUTER_API_KEY` | Optional | Enables AI narrative + chat (`deepseek/deepseek-chat-v3-0324:free` by default) |
| `OPENROUTER_MODEL` | Optional | Override the default model |
| `SCE_UPLOAD_MAX_MB` | Optional | Max file size in MB (default 50) |

## File requirements

- Primavera P6 `.XER` format
- Must contain `PROJECT` and `TASK` tables
- Max 50 MB per file

## Run tests

```bash
python -m pytest tests/ -v
```
