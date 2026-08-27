# Operations KPI Dashboard

Custom browser dashboard for aggregate OPS KPIs. The repository contains no operational figures by default.

## Accepted upload

Upload an `.xlsx` file containing a worksheet named `Dashboard Export`. Required columns are `Year`, `Month`, `Branch`, `Group`, `Measure`, and `Value`; optional columns are `Target`, `Unit`, and `UpdatedAt`.

Supported groups: Unit Savings, Monthly Unit Progress, Hours Worked Breakdown, Outlook, Agency vs Permanent, Budget, and Gross Profit. The Analysis tab uses weekly snapshots for Unit Savings, Monthly Unit Progress and Hours Worked; completed monthly periods for Agency vs Permanent and Gross Profit; and YTD aggregates through the selected period.

## Privacy

Do not commit the source OPS workbook. GitHub Pages is public unless the organisation enables private Pages through GitHub Enterprise Cloud. Only aggregated export rows should ever be published.

## Planned twice-daily refresh

Power Automate refreshes the Power BI-connected workbook, validates the `Dashboard Export` table, and sends a GitHub repository-dispatch event. GitHub Actions deploys the validated aggregate JSON. Keep `testMode` enabled while Pages is public so only the Demo Branch dataset is published. Real OPS figures must remain unpublished until the dashboard is moved to restricted SharePoint hosting.

See `power-automate/SETUP.md` and `office-scripts/refresh-and-export.ts`. The flow must run under a Microsoft 365 account that can refresh the Power BI connection and access the SharePoint/OneDrive Outlook workbook. The Unit Count Forecast workbook is intentionally excluded.

## Branch and geography handling

Selectable branch names are produced by the Office Script from the approved branch and geography-code mappings. Gross Profit shown under **All branches** includes every geography code, including central, project and other codes that are deliberately excluded from the individual branch selector. Any new, unknown or unmapped geography-code check returned by the refresh payload is highlighted in the dashboard's Refresh checks section for review.
