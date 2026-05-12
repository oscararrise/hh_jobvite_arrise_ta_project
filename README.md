# HH.ru to Jobvite Automation

## Overview

This project automates the transfer of qualified candidates from HH.ru into Jobvite.

The automation retrieves active HH.ru vacancies, identifies the related Jobvite requisition ID from each vacancy description, filters candidates based on their test score, extracts and transforms candidate data, downloads the candidate resume, and creates the candidate/application record in Jobvite.

The process also generates an Excel report with execution details, requisition summaries, and execution-level metrics.

## Main Features

- Retrieves active vacancies from HH.ru by manager.
- Extracts Jobvite requisition IDs from HH.ru vacancy descriptions.
- Retrieves candidates from HH.ru negotiations.
- Filters candidates by minimum test score.
- Retrieves full resume data from HH.ru.
- Extracts candidate contact information.
- Converts Cyrillic names and work history to Latin characters.
- Downloads candidate CV in PDF format.
- Sends candidate data and resume to Jobvite.
- Creates candidates with application status set to `New`.
- Generates an Excel report with detailed and aggregated execution results.

## Project Structure

```text
.
├── config.py
├── main.py
├── statics.py
├── requirements.txt
├── .gitignore
├── README.md
└── reports/
```

## File Description

### `main.py`

Main execution file.

It controls the complete automation flow:

1. Loads the HH.ru access token.
2. Retrieves managers from HH.ru.
3. Retrieves active vacancies for each manager.
4. Reads vacancy details.
5. Extracts the Jobvite requisition ID from the vacancy description.
6. Retrieves qualified candidates.
7. Creates candidates in Jobvite.
8. Saves the execution report.

### `config.py`

Contains the core helper functions used by the automation, including:

- HH.ru API calls.
- Jobvite API calls.
- Candidate payload creation.
- Resume download and base64 conversion.
- Candidate data transformation.
- Excel report generation.

### `statics.py`

Loads environment variables from the `.env` file and exposes configuration values used by the application.

This file should not contain hardcoded credentials.

### `.env`

Local environment file used to store secrets, API keys, access tokens, URLs, and credentials.

This file must never be committed to Git.

### `requirements.txt`

Contains the Python dependencies required to run the project.

### `reports/`

Directory where the Excel execution report is generated.

Generated reports are excluded from Git by default.

## Environment Variables

Create a `.env` file in the root directory with the following structure:

```env
# =========================
# JOBVITE
# =========================
JOBVITE_API_KEY=
JOBVITE_API_SECRET=
JOBVITE_COMPANY_ID=

JOBVITE_JOB_URL=https://api.jobvite.com/api/v2/job
JOBVITE_CANDIDATE_URL=https://api.jobvite.com/api/v2/candidate


# =========================
# HH.RU / HH.KZ
# =========================
HH_ACCESS_TOKEN=
HH_USER_AGENT=

HH_EMPLOYER_ID=
HH_CLIENT_ID=
HH_CLIENT_SECRET=
HH_REDIRECT_URI=http://localhost:8080/callback

HH_AUTH_URL=https://hh.ru/oauth/authorize
HH_TOKEN_URL=https://hh.ru/oauth/token
HH_API_URL=https://api.hh.ru


# =========================
# NAUKRI / GULF
# =========================
NAUKRI_SECRET_KEY=
NAUKRI_ACCESS_KEY=
```

## Security Notes

The `.env` file contains sensitive credentials and must not be committed to the repository.

Before pushing changes to a remote repository, always verify that `.env` is ignored by Git:

```bash
git status
```

The `.env` file should not appear in the list of tracked or untracked files.

The following files and folders are ignored by default:

```text
.env
*.env
__pycache__/
*.pyc
.vscode/
reports/*.xlsx
reports/*.xls
*.log
```

## Requirements

This project requires Python 3.10 or higher.

Install the dependencies using:

```bash
pip install -r requirements.txt
```

Required packages:

```text
requests
python-dotenv
cyrtranslit
openpyxl
```

## How to Run

From the root directory of the project, run:

```bash
python main.py
```

The automation will:

1. Load environment variables.
2. Connect to HH.ru.
3. Retrieve managers and vacancies.
4. Process candidates.
5. Create candidates in Jobvite.
6. Generate or update the Excel report.

## Report Output

The execution report is generated in:

```text
reports/jobvite_hhru_publication_report.xlsx
```

The workbook contains the following sheets:

### `Posting Details`

Detailed list of successfully posted candidates.

### `Summary Requisitions`

Aggregated summary by Jobvite requisition and HH.ru vacancy.

### `Summary Execution`

Execution-level metrics per automation run.

## Candidate Filtering Logic

Candidates are retrieved from HH.ru negotiations and filtered by test score.

The default minimum score is configured inside the `get_top_candidates` function:

```python
min_score: int = 84
```

Only candidates with a score equal to or greater than this value are processed.

## Requisition Mapping Logic

The automation extracts the Jobvite requisition ID from the HH.ru vacancy description using this pattern:

```text
REQ <number>
REQ: <number>
REQ# <number>
REQ- <number>
```

Example:

```text
REQ 6181
```

If no requisition ID is found in the vacancy description, the vacancy is skipped.

## Jobvite Candidate Creation

Each candidate is created in Jobvite with the following application information:

```text
workflowState: New
source: HH
sourceType: JobBoard
```

The candidate resume is attached as a PDF using base64 encoding.

## Local Development Workflow

Recommended Git workflow:

```bash
git status
git add .
git commit -m "Describe your change"
git push
```

Before each commit, verify that sensitive files are not included:

```bash
git status
```

## Deployment Notes

For deployment on a virtual machine, the following steps are required:

1. Clone the repository.
2. Create the `.env` file manually on the server.
3. Install Python dependencies.
4. Run the script manually to validate connectivity.
5. Configure a scheduled execution using cron or a process manager.

Example clone command:

```bash
git clone <repository-url>
```

Example dependency installation:

```bash
pip install -r requirements.txt
```

Example execution:

```bash
python main.py
```

## Suggested Production Schedule

The automation can be scheduled to run once per day.

Recommended execution time:

```text
07:30 CET
```

## Important Operational Considerations

- HH.ru access tokens may expire and need to be refreshed.
- Jobvite API failures should be monitored.
- Duplicate candidate handling should be validated before production rollout.
- Generated Excel reports should not be committed to Git.
- Credentials must be managed through environment variables only.
- Production credentials should be rotated before deployment.
- Logs should be reviewed after each scheduled execution during the first production days.
- Token refresh logic should be reviewed before production if the process is expected to run unattended.
- Candidate duplicate validation should be confirmed with Jobvite before enabling the process in production.

## Troubleshooting

### Git does not detect changes

Run:

```bash
git status
```

### Environment variable is missing

Check that the `.env` file exists in the root directory and contains the required key.

### Python cannot find `dotenv`

Install dependencies:

```bash
pip install -r requirements.txt
```

### HH.ru returns unauthorized response

Validate:

- `HH_ACCESS_TOKEN`
- OAuth token expiration
- Employer ID
- User permissions
- HH.ru API access

### Jobvite rejects candidate creation

Validate:

- Jobvite API key
- Jobvite secret
- Company ID
- Requisition ID
- Candidate payload format
- Resume attachment format

### Excel report is not generated

Validate:

- The `reports/` directory can be created by the script.
- The user running the script has write permissions.
- The Excel file is not open while the script is running.

## Repository Status

Initial version prepared for deployment and production validation.

## Maintainer

Oscar Jimenez  
Talent Analytics and Technology  
ARRISE