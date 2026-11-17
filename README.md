# Participant-Information Management System — HB-GTIs Trial

An R/Shiny application used to manage participant schedules, assessments, and
data for the HB-GTIs trial: a randomized, double-blind, sham-controlled study
of home-based, globus pallidus internus (GPi)-targeted transcranial temporal
interference stimulation (TIs) for motor symptoms in Parkinson's disease.

This repository contains the source code and automated tests only. It does not
contain participant-level data, credentials, or database files.

## Features

- Participant records keyed by a unique study identification code
- Visit scheduling generated from the two baseline visit dates
- Electronic capture of clinical scales, MDS-UPDRS Part III, and the caregiver diary
- Automated scoring of totals, subscales, and change from the previous assessment
- Data-integrity safeguards at the point of entry
- Automated regression tests for scheduling and batch completion

## Repository structure

```
app.R
tests/
  test_schedule_without_midpoint.R
  test_daily_home_completion.R
README.md
LICENSE
```

## Requirements

- R (>= [4.2])
- R packages: `shiny`, `DBI`, `RSQLite`, `jsonlite`, `[其他包]`
- Optional: `RPostgres` (only if using PostgreSQL)

Install the required packages:

```r
install.packages(c("shiny", "DBI", "RSQLite", "jsonlite"))
```

## Configuration

The application reads the following environment variables:

| Variable | Default | Description |
|---|---|---|
| `DB_DRIVER` | `sqlite` | Database driver; set to `postgres` to use PostgreSQL |
| `SQLITE_DB_FILE` | `[path]` | Path to the SQLite database file |

Example (Linux/macOS):

```bash
export DB_DRIVER=sqlite
export SQLITE_DB_FILE=/path/to/trial.sqlite
```

Example (Windows PowerShell):

```powershell
$env:DB_DRIVER="sqlite"
$env:SQLITE_DB_FILE="C:\path\to\trial.sqlite"
```

## Running the application

1. Install the required packages (see Requirements).
2. Set the environment variables (see Configuration).
3. Launch the app:

```r
shiny::runApp("app.R")
```

Alternatively, open `app.R` in RStudio and click **Run App**.

## Data model

The system uses five tables:

- `subjects` — one row per participant (study ID, screening/baseline dates,
  contact details, withdrawal status and reason).
- `visits` — the event schedule and execution status (dates, operator, timestamps).
- `scale_records` — one row per completed instrument (scale type, participant,
  date, visit, total score, item-level responses as JSON).
- `updrs3_records` — MDS-UPDRS Part III assessments.
- `family_diary_records` — caregiver diary entries.

## Data privacy

This repository contains no participant-level data, no personal identifiers,
no database files, and no credentials. All example configurations are
placeholders.

## License

This project is released under the [MIT License](LICENSE).

## Contact

[Louis]，[Shanghai University of Sport] — [luzijun@snnu.edu.cn]
