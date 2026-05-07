# AWS Monthly Cost — Automatic Billing Report

A Python CLI tool that fetches billing data from multiple AWS accounts using the **Cost Explorer API** and generates a professional, multi-sheet Excel report — complete with tax breakdowns, month-over-month comparisons, 5-month history, service-level analysis, and embedded pie charts.

---

## Table of Contents

- [Features](#features)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [AWS IAM Permissions](#aws-iam-permissions)
- [Configuration](#configuration)
  - [Option 1 — Environment Variable (Recommended)](#option-1--environment-variable-recommended)
  - [Option 2 — YAML Configuration File (Fallback)](#option-2--yaml-configuration-file-fallback)
- [Running the Application](#running-the-application)
- [Output](#output)
- [Excel Report Sheets](#excel-report-sheets)
- [Troubleshooting](#troubleshooting)
- [Dependencies](#dependencies)

---

## Features

| Feature | Description |
|---|---|
| **Multi-Account** | Processes any number of AWS accounts in a single run |
| **No `~/.aws/credentials` Required** | Supports inline access keys via environment variable — no local credential file needed |
| **Tax Segregation** | Separates costs into **with tax** and **without tax** (excludes Tax, Credits, Refunds) |
| **Month-over-Month Comparison** | Compares the last two completed months and flags Increased / Decreased / Stable |
| **Primary Reason Detection** | Identifies the top 3 services responsible for cost changes |
| **5-Month History** | Shows billing trend for the last 5 completed months per account |
| **Service-Level Breakdown** | Lists top 10 services by cost with percentage share and embedded pie charts |
| **Professional Excel Output** | Styled workbook with color-coded headers, alternating rows, conditional formatting, and frozen panes |

---

## Project Structure

```
aws-monthly-cost-automatic-billing/
│
├── main.py                      # Entry point — orchestrates everything
├── config/
│   └── accounts.yaml            # Fallback account configuration (YAML)
├── src/
│   ├── config_loader.py         # Loads accounts from env var or YAML
│   ├── aws_cost.py              # Fetches monthly & historical costs from Cost Explorer
│   ├── service_analysis.py      # Per-service cost breakdown & change analysis
│   └── report_generator.py      # Generates the styled multi-sheet Excel report
├── output/                      # Generated reports saved here
│   └── aws_report_april_2026.xlsx
└── README.md
```

---

## Prerequisites

- **Python** 3.8 or higher
- Valid **AWS credentials** for each target account (access key + secret key, OR configured AWS CLI profiles)
- IAM permissions to access the **Cost Explorer API** (see [AWS IAM Permissions](#aws-iam-permissions))

---

## Installation

1. Clone this repository:

   ```bash
   git clone <repo-url>
   cd aws-monthly-cost-automatic-billing
   ```

2. Install dependencies:

   ```bash
   pip install boto3 openpyxl pyyaml python-dateutil
   ```

---

## AWS IAM Permissions

The IAM user or role associated with each account's credentials **must** have the following permission:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "ce:GetCostAndUsage"
      ],
      "Resource": "*"
    }
  ]
}
```

> **Note:** Cost Explorer must also be **enabled** in each AWS account. The first time you access Cost Explorer in a new account, it may take up to 24 hours before data becomes available.

---

## Configuration

The application loads account information in the following priority order:

1. **`AWS_ACCOUNTS` environment variable** — checked first
2. **`config/accounts.yaml` file** — used as fallback if the env variable is not set

---

### Option 1 — Environment Variable (Recommended)

Set the **`AWS_ACCOUNTS`** environment variable. This is the recommended method, especially if you **don't want to use `~/.aws/credentials`**.

#### Variable Name

```
AWS_ACCOUNTS
```

#### Format

A **comma-separated** list of accounts. Each account uses **colon-separated** fields.

Two formats are supported — you can even **mix them** in the same string:

| Format | Fields | When to Use |
|---|---|---|
| **Direct Credentials** | `Name:AccountID:AccessKeyID:SecretAccessKey` | ✅ When you don't want to use `~/.aws/credentials` |
| **AWS Profile** | `Name:AccountID:ProfileName` | When credentials are already in `~/.aws/credentials` |

#### Using Direct Credentials (No `~/.aws/credentials` Needed)

```bash
export AWS_ACCOUNTS="MyAccount1:111111111111:AKIAIOSFODNN7EXAMPLE:wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY,MyAccount2:222222222222:AKIAI44QH8DHBEXAMPLE:je7MtGbClwBF/2Zp9Uo7OiEXAMPLEKEYSECRET"
```

**Full example with multiple accounts:**

```bash
export AWS_ACCOUNTS="Production:111111111111:AKIAIOSFODNN7EXAMPLE:wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY,Staging:222222222222:AKIAI44QH8DHBEXAMPLE:je7MtGbClwBF/2Zp9Uo7OiEXAMPLEKEYSECRET,Development:333333333333:AKIDEXAMPLE3RDACCTKY:zBqMF8d3PLaYIU0x7IEXAMPLEKEYFORDEV12"
```

**Field breakdown:**

```
       Name        : Account ID   : Access Key ID        : Secret Access Key
       ↓              ↓              ↓                       ↓
    Production  :  111111111111  :  AKIAIOSFOD...  :  wJalrXUtnFEM...
```

#### Using AWS Profiles

If your credentials are already in `~/.aws/credentials`:

```bash
export AWS_ACCOUNTS="Staging:222222222222:staging-profile,Production:111111111111:production-profile"
```

#### Windows (PowerShell)

```powershell
$env:AWS_ACCOUNTS="Production:111111111111:AKIA...:wJal...,Staging:222222222222:AKIA...:je7M..."
```

#### Making It Persistent

**Mac/Linux** — Add to your shell config (`~/.zshrc` or `~/.bashrc`):

```bash
echo 'export AWS_ACCOUNTS="..."' >> ~/.zshrc
source ~/.zshrc
```

> ⚠️ **Security Warning:** Avoid storing access keys in plain text in shell config files. For better security, consider using a `.env` file with [direnv](https://direnv.net/) or a secrets manager.

---

### Option 2 — YAML Configuration File (Fallback)

If `AWS_ACCOUNTS` is **not set**, the application reads from `config/accounts.yaml`.

This method only supports **AWS profile names** (requires `~/.aws/credentials`):

```yaml
accounts:
  - name: Production
    id: "111111111111"
    profile: production-profile

  - name: Staging
    id: "222222222222"
    profile: staging-profile

  - name: Development
    id: "333333333333"
    profile: dev-profile
```

Each entry requires:

| Field | Description | Example |
|---|---|---|
| `name` | Display name for the account (used in the report) | `Production` |
| `id` | AWS Account ID (12 digits) | `"111111111111"` |
| `profile` | AWS CLI profile name from `~/.aws/credentials` | `production-profile` |

---

## Running the Application

```bash
python main.py
```

### What Happens

1. **Loads accounts** — from `AWS_ACCOUNTS` env variable or `config/accounts.yaml`
2. **Iterates through each account** — creates a boto3 session per account
3. **Fetches billing data** — for the last two completed months:
   - Total cost (with and without tax)
   - Per-service cost breakdown
   - Primary reasons for cost change (top 3 services)
   - 5-month historical cost trend
4. **Generates Excel report** — saved to the `output/` directory
5. **Prints progress** — to the terminal with ✔/❌ indicators

### Expected Terminal Output

```
✔ Loaded 3 account(s) from AWS_ACCOUNTS env variable

----------------------------
Processing account: Production
----------------------------
  Fetching costs with tax...
  Fetching costs without tax...
✔ Success: Production

----------------------------
Processing account: Staging
----------------------------
  Fetching costs with tax...
  Fetching costs without tax...
✔ Success: Staging

...

Generating Excel report...
✔ Report created: output/aws_report_april_2026.xlsx
```

---

## Output

Reports are saved in the `output/` directory with the naming convention:

```
output/aws_report_{month}_{year}.xlsx
```

Example: `output/aws_report_april_2026.xlsx`

---

## Excel Report Sheets

The generated workbook contains **4 professionally styled sheets**:

### Sheet 1 — Account Cost Summary

| Column | Description |
|---|---|
| SN | Serial number |
| Account Name | AWS account display name |
| Cost (Excl. Tax) | Current month cost excluding tax, credits, refunds |
| Tax | Tax amount for the current month |
| Total (Incl. Tax) | Full Cost Explorer cost including everything |

Includes a **total row** at the bottom summing all accounts.

### Sheet 2 — Monthly Cost Observation

| Column | Description |
|---|---|
| Account Name | AWS account display name |
| Account ID | 12-digit AWS account ID |
| Previous Month | Cost of the month before last (excluding tax) |
| Current Month | Cost of the last completed month (excluding tax) |
| Change | Dollar difference between the two months |
| Status | `Increased` (🔴), `Decreased` (🟢), or `Stable` (🟡) — color-coded |
| Primary Reason | Top 3 services driving the cost change |

### Sheet 3 — Last 5 Months Billing History

For **each account**, a separate table showing:

| Column | Description |
|---|---|
| Month | Month name (e.g., Dec 2025) |
| Cost (Excl. Tax) | Monthly cost excluding tax |
| Tax | Tax amount |
| Total (Incl. Tax) | Full monthly cost |

### Sheet 4 — Service Level Cost Breakdown

For **each account**, a table and **embedded pie chart** showing:

| Column | Description |
|---|---|
| Service | AWS service name (e.g., Amazon EC2, Amazon S3) |
| Cost (USD) | Service cost for the current month |
| % | Percentage of total account cost |

- Shows the **top 10 services**; remaining services are grouped as **"Other"**
- Pie chart displays the **top 5 services** with percentage labels

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `AccessDeniedException` | Missing IAM permissions | Add `ce:GetCostAndUsage` permission to the IAM user/role |
| `ProfileNotFound` | AWS profile not in `~/.aws/credentials` | Use direct credentials in `AWS_ACCOUNTS` env var instead, or add the profile to `~/.aws/credentials` |
| `ValueError: AWS_ACCOUNTS env variable is set but contains no valid entries` | Malformed `AWS_ACCOUNTS` value | Ensure format is `Name:ID:AccessKey:SecretKey` (4 parts) or `Name:ID:Profile` (3 parts), comma-separated |
| `No module named 'boto3'` | Dependencies not installed | Run `pip install boto3 openpyxl pyyaml python-dateutil` |
| `FileNotFoundError: config/accounts.yaml` | No env var set and YAML file missing | Either set `AWS_ACCOUNTS` env variable or create `config/accounts.yaml` |
| `Cost Explorer data not available` | Cost Explorer not enabled or new account | Enable Cost Explorer in the AWS Console; data takes up to 24 hours to appear |

---

## Dependencies

| Package | Purpose |
|---|---|
| `boto3` | AWS SDK — interacts with Cost Explorer API |
| `openpyxl` | Creates and styles the Excel workbook |
| `pyyaml` | Parses the YAML configuration file |
| `python-dateutil` | Date math for month calculations |

Install all:

```bash
pip install boto3 openpyxl pyyaml python-dateutil
```

---

## License

Internal use only.
