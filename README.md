# Marlu Weekly Timesheet Automation

## Overview

This project is an internal Azure Function automation system used for weekly timesheet processing and roster hour reconciliation.

The automation is designed to:

* Pull OPMS timesheet data
* Match against SharePoint roster records
* Calculate effective worked hours
* Generate weekly Excel reports
* Support automated email distribution
* Prepare future scheduled Azure Function execution

---

## Main Features

### Weekly Timesheet Processing

* Aggregates OPMS daily timesheet data
* Calculates employee worked hours
* Handles roster matching logic

### Roster Reconciliation

* Matches:

  * OPMS Employee ID
  * SharePoint roster dates
* Updates final hours used for payroll/reporting

### Excel Report Generation

Automatically exports:

* Weekly timesheet reports
* Gap summaries
* Unmatched records
* Roster summaries

### Azure Function Ready

Designed for:

* Azure Functions
* Scheduled execution
* GitHub CI/CD deployment

---

## Project Structure

```text
Timesheethour_main.py        # Main execution script
Timesheethours1.py           # Processing logic
Timesheethours2.py           # Data aggregation logic
TimesheethourFormat.py       # Excel formatting/export
TimesheethourEmail.py        # Email functionality
.gitignore                   # Git exclusions
```

---

## Environment Variables

The following secrets/configuration are stored locally using `.env` and are excluded from GitHub:

* TENANT_ID
* CLIENT_ID
* CLIENT_SECRET
* SHAREPOINT_HOST
* SITE_NAME

---

## Deployment

This project is intended to be deployed using:

* Azure Functions
* GitHub Actions
* Python 3.11

---

## Notes

This repository excludes:

* Local secrets
* Excel exports
* Temporary checkpoint files
* Internal web application files
* HTML templates

For internal use only.
