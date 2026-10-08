# Lead Quality Checker — corrected replacement

## Put this on GitHub / Railway
1. Unzip and upload the contents of this folder to your existing repository root. Replace your old app file with `app.py`, or change your Railway start command to use `app.py`.
2. Include `checker.py`, `requirements.txt`, and `Procfile` alongside it.
3. Railway start command: `streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT --server.headless=true`.
4. Locally: `pip install -r requirements.txt`, then `streamlit run app.py`.
5. Upload the master and optional picklist. Check the visible column mappings before running. Download results.

Python 3.10+ required. Deployment has not been performed. Existing Railway settings can override the Procfile. No API keys required.

## What changed
- Email/company, website/company and email/website checks are independent.
- Exact normalized column aliases plus manual mapping replace unsafe substring guessing.
- Domain parsing understands subdomains and public suffixes, using the bundled suffix list without network fetching.
- No generic government-domain pass or permissive abbreviation-subsequence pass.
- Missing or unverified checks produce REVIEW, never a silent pass.
- Phone country is checked against the supplied country; national numbers no longer default to US.
- Instruction rows can be skipped. Original sheet names, fields and other sheets are retained.
- Results use QA_ columns. On reruns those columns are overwritten, not left stale. Old Match_/Company_Domain_ columns from the previous tool are preserved but are not used: rely only on QA_ results.
- Source uploads/settings changing clears the previous download.
- The unused Commercial/Wholesale selector and inferred seniority were removed; campaign-specific rules were not present in the supplied code. Picklists still check supplied seniority/job level.

## Verified relationship file
Upload a CSV/XLSX with exactly these required headers:
`company,domain,relationship,evidence_url,checked_date`

Allowed relationship values:
- `official`, `trading_name`: green, verified company/domain relationship.
- `parent`, `subsidiary`, `acquisition`, `rebrand`: amber, connected but applicability needs review.
- `unrelated`: red, manually researched conflict.

Dates must be YYYY-MM-DD. Evidence older than 180 days requires refresh. One row per company/domain pair; duplicate pairs are rejected. Company matching uses the normalized full company name; add another row if you need a different name alias. The supplied template is intentionally empty: no relationships have been invented or independently verified. Relationship data is trusted operator input; a source URL is required but not fetched or validated by this app. Store the CSV securely yourself and upload it again on future runs.

Name similarity alone is yellow. Unknown relationships are yellow, not red. Matching email/website domains alone never verifies the company or person. A green relationship result does not verify employment, mailbox ownership or deliverability. A connected-domain finding never automatically approves the lead.

## Research / AI scope
This release does NOT perform web searches, automatic AI research, current-employer verification or email discovery. It supports recorded, evidence-backed relationships and identifies what needs research. Automatic research requires a separately configured search/model service and a review workflow. No lead information is sent to an external AI service by this code.

## Workbook handling
The input format is XLSX, with headers in row 1. Picklists use one column per field. Blank picklist cells are ignored. Missing picklists remain visibly unchecked. All supplied criteria must pass for overall checks-passed status; employment/email verification remains separate.

The template-row heuristic skips rows containing at least three standard instruction markers. Turn it off for nonstandard data. Employee ranges stored by Excel as dates must be corrected at source; this tool cannot reliably recover the intended ranges. Existing Excel charts/unsupported advanced objects may not round-trip perfectly through openpyxl; keep your original workbook.

## Verification
`python -m unittest -v`
Five regression tests cover false header detection, public-suffix parsing, relationship flags, independent email checks, repeated processing, and phone-country mismatch.
Both supplied four-lead workbooks were processed successfully: all four leads require review. The master template row was skipped. Personal sample records are not included in this package.
