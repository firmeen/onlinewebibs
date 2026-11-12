# OIC XMLA Collector — Excel + SQL Server

Reads 5 insurance cubes: **AH, FI, MR, MI, MV**. Source is an OIC XMLA `Execute` POST with the exact MDX/Catalog from the supplied browser capture. Destination: Excel `.xlsx` (created with Python standard library, no separate Excel package) and/or SQL Server via **python-tds (`pytds`) only**. `pyodbc` is never used.

## Install (PowerShell)

```powershell
python -m pip install -r requirements.txt
```

## Live Excel in one command (using local `.env`)

Create `.env` in this folder (it is gitignored). Paste the `OIC_BEARER_TOKEN` and
`OIC_COOKIE` from a fresh, authorized browser session. The collector loads `.env`
automatically; you do **not** need to set PowerShell environment variables.

```powershell
Copy-Item .env.example .env
notepad .env
python oic_collector.py --mode live --output excel
```

Excel is saved to `output/oic_insurance_YYYYMMDD_HHMMSS.xlsx`.
Use `--output both` if you also configure SQL credentials in `.env`.

## Test from the captured file (no network, no token, Excel only)

```powershell
python oic_collector.py --mode offline --source-file "C:\path\to\capture.txt" --output excel
```

## Connect to SQL Server

```powershell
$env:SQL_SERVER = "YOUR_SQL_SERVER"
$env:SQL_PORT = "1433"
$env:SQL_DATABASE = "YOUR_DATABASE"
$env:SQL_USER = "YOUR_SQL_USER"
$env:SQL_PASSWORD = Read-Host "SQL password"
```

Then import the *captured* XML data into **both outputs**:

```powershell
python oic_collector.py --mode offline --source-file "C:\path\to\capture.txt" --output both
```

## Request new live data

Get a **fresh** Bearer token from your own authorized browser session. Don't paste tokens into code or Git. If the server also requires an authorized session cookie, set `OIC_COOKIE` in your environment; no authentication bypass is attempted.

```powershell
$env:OIC_BEARER_TOKEN = Read-Host "Authorized OIC Bearer token"
# $env:OIC_COOKIE = Read-Host "OIC cookie (only if required)"
python oic_collector.py --mode live --output both
```

To download only AH and MV:

```powershell
python oic_collector.py --mode live --codes AH MV --output both
```

- Default SQL table: `dbo.OIC_INSURANCE_METRICS` (may require `CREATE TABLE` permission the first time).
- Each category is **replaced in a transaction**, not appended, preventing duplicate records between runs. Other categories remain unchanged.
- `aggregation_level` separates `TOTAL`, `SUBTOTAL`, `DETAIL` to prevent double-counting.
- `metric_raw` retains source precision; Excel's `metric_value` is numeric but may be subject to Excel's 15 significant-digit limit. SQL stores `DECIMAL(38,12)` alongside the exact raw value.
- MV Request's **webpage Class=MO** is intentionally different from its XMLA cube/catalog suffix **MV**, matching the capture.
- The sample capture includes historical aggregate data and metadata, not individual policy records. Access permissions, valid tokens and server responses must be tested in your environment.
- **Security:** The original capture contains a Bearer token and cookies. Do not upload or commit it. Revoke/rotate exposed session credentials according to your organization's procedure.