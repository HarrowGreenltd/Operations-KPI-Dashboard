# Operations KPI Dashboard

Custom browser dashboard for aggregate OPS KPIs. The repository contains no operational figures by default.

## Accepted upload

Upload an `.xlsx` file containing a worksheet named `Dashboard Export`. Required columns are `Year`, `Month`, `Branch`, `Group`, `Measure`, and `Value`; optional columns are `Target`, `Unit`, and `UpdatedAt`.

Supported groups: Unit Savings, Monthly Unit Progress, Hours Worked Breakdown, Outlook, Agency vs Permanent, and Budget.

## Privacy

Do not commit the source OPS workbook. GitHub Pages is public unless the organisation enables private Pages through GitHub Enterprise Cloud. Only aggregated export rows should ever be published.

## Planned twice-daily refresh

Power Automate refreshes the Power BI-connected workbook, validates the `Dashboard Export` table, and sends a GitHub repository-dispatch event. GitHub Actions deploys the validated aggregate JSON. Keep the flow disabled while Pages is public and real OPS data is being used.

See `power-automate/SETUP.md` and `office-scripts/refresh-and-export.ts`. The flow must run under a Microsoft 365 account that can refresh the Power BI connection and access the SharePoint/OneDrive Outlook workbook. The Unit Count Forecast workbook is intentionally excluded.
