# DEV Infrastructure Monitoring Automation

## Files

- `monitoring.py` - main monitoring program
- `urls.txt` - project-wise URL list
- `mongodb.txt` - MongoDB backup configuration
- `rabbitmq.txt` - RabbitMQ message-store configuration
- `cassandra.txt` - Cassandra configuration
- `credentials.txt` - local SSH username/password
- `requirements.txt` - Python dependencies
- `Jenkinsfile` - Jenkins pipeline
- `.gitignore` - prevents credentials/report output from being committed

## Local Windows test

Open PowerShell in this folder:

```powershell
py -m pip install -r requirements.txt
py .\monitoring.py
```

Report:

```text
monitoring_output\monitoring_report.xlsx
```

## SSH

The script uses password authentication:

```text
username=controller
password=nav#123
```

For RabbitMQ, the script uses:

```bash
sudo -S
```

and sends the sudo password through the SSH command stdin because the `controller`
user currently requires a sudo password.

## MongoDB backup date

The script always checks yesterday.

Example:

```text
Today:      22-Sep-2026
Backup:     21Sep2026
```

India is checked daily.

US is checked only when yesterday was Sunday.

A valid existing backup directory is considered successful even if its size is 0.

## URL behavior

For every URL:

1. Try configured HTTP/HTTPS URL.
2. If it fails, swap HTTP <-> HTTPS.
3. If the alternate works, report the working URL.
4. If both fail, report `No`.

HTTPS certificate verification is disabled because the current environment contains
internal/self-signed HTTPS endpoints.

## Excel

The report contains:

- URL Monitoring
- MongoDB Backup
- RabbitMQ Message Store
- Cassandra

Projects in the URL section are vertically merged, so the project name appears once
for consecutive services instead of repeating on every row.

Status colors:

- YES = green
- NO = red
- N/A = yellow

## Important

Do not commit `credentials.txt`.

For Jenkins production use, put the SSH username/password into Jenkins Credentials
and inject them as environment variables or a protected credentials file.
