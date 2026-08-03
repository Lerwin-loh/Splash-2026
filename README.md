# Splash Mentoring Allocation

A small Flask web application for allocating event applicants into two main groups and two mentoring subdomains. It is designed for local use or Vercel deployment, with no database, authentication, background jobs, or permanent file storage.

## Features

- Upload an `.xlsx` registration workbook.
- Inspect worksheets, detected columns, applicant count, and the first five rows.
- Confirm or change the three preference columns.
- Configure maximum applicants and mentor counts by subdomain.
- Allocate applicants into Group 1 and Group 2 by original registration order.
- Assign each applicant to two distinct mentoring subdomains whenever feasible.
- Display operational analytics in the browser.
- Download a formatted Excel workbook with allocated applicants and a simple allocation summary.

## Technology Stack

- Python
- Flask
- Pandas
- OpenPyXL
- HTML, CSS, and plain JavaScript
- Pytest

## Supported Excel Format

The workbook should contain one applicant per row. The application preserves all original columns and appends:

- Applicant Number
- Grouping
- Event Sequence
- Subdomain Allocation 1
- Subdomain Allocation 2
- Allocation Status
- Attending

The three preference columns can use names such as `1st Option`, `First Choice`, `2ndOption`, `Second Choice`, `3rd Option`, or `Third Choice`. The frontend lets you confirm the detected mappings.

Supported subdomains are:

- Cybersecurity
- Quantum Computing
- Cloud Computing
- Artificial Intelligence

Common variations such as `AI`, `Cyber security`, and different capitalisation are normalised.

## Allocation Rules

Applicants are numbered by their original non-blank Excel row order. Odd-numbered applicants go to Group 1 and even-numbered applicants go to Group 2.

Group 1 attends:

`Panel Discussion -> Mentoring`

Group 2 attends:

`Mentoring -> Panel Discussion`

Each applicant receives two different mentoring subdomains. The app tries ranked preferences first, then uses a deterministic fallback based on the greatest remaining capacity percentage.

Mentor names are not needed. Mentor counts only define capacity:

`subdomain capacity = mentors in subdomain x participants per mentor`

The starting participants-per-mentor limit is 5. If a group cannot be fully allocated, the app increases the limit by 1 and retries. Group 1 and Group 2 are calculated independently because the same mentor capacity is reused across the two mentoring sessions.

## Local Setup

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:5000/
```

## Run Tests

```bash
pytest
```

## Vercel Deployment

This project includes `vercel.json` and uses `api/index.py` as the Flask entrypoint.

Typical deployment:

```bash
vercel
```

Then follow the Vercel prompts. No database or environment variables are required.

## File Privacy Behaviour

Uploaded files are processed only for the current request. The application does not retain uploaded workbooks, does not use a database, and does not send applicant data to external APIs. Malformed files and unsupported file types are rejected with user-friendly errors.

## Sample Workbook

`sample_registration_200.xlsx` contains 200 fake registrations for testing the upload and allocation flow. It includes valid preferences, spelling variations, duplicate preferences, blank preferences, invalid preferences, leading-zero phone and membership numbers, alternate option-column headers, a worksheet with a blank row, and an empty template worksheet.

`sample_registration.xlsx` is the original smaller 20-row sample.

## Known Limitations

- Only `.xlsx` files are supported.
- The upload size limit is 8 MB.
- The app is intended for roughly 400 applicants, not very large conferences.
- Individual mentor assignment is intentionally not supported.
- Analytics are operational summaries, not a detailed preference-satisfaction report.
