# Twice-daily refresh with weekly analysis

This design keeps the calculation workbook in SharePoint/OneDrive and publishes only the `Dashboard Export` aggregate table.

1. Store the OPS workbook and `HG Outlook - Mar 26` in SharePoint/OneDrive locations accessible to the flow owner. Keep the Outlook link stable.
2. Add the Office Script in `office-scripts/refresh-and-export.ts` to Excel for the web.
3. Keep the scheduled cloud flow Recurrence trigger at every 12 hours in the `Europe/London` time zone.
4. Add **Excel Online (Business) → Run script**, select the OPS workbook, and run `refresh-and-export`.
5. Add a GitHub **Repository dispatch** call for `HarrowGreenltd/Operations-KPI-Dashboard` with event type `ops-kpi-refresh`. Send the script result as the `client_payload` body.
6. Test with `isTestData: true` while the site is public. Do not publish real OPS figures until private GitHub Pages is available and access is configured.
7. Enable the flow only after the workbook refresh and dashboard checks both pass. The first successful run creates the current week's provisional baseline; prior-week commentary becomes available when a second weekly period exists.

The Office Script creates and maintains an `OpsDashboardHistory` table in a `Dashboard History` worksheet. It retains aggregate Monday-to-Sunday weekly snapshots for Unit Savings, Monthly Unit Progress and Hours Worked; no raw job or customer data is added. Each twice-daily refresh replaces the current week's provisional values instead of creating another comparison point. When the next week begins, the previous week's last values remain fixed. This avoids granting the GitHub workflow permission to write back to the repository.

The Analysis tab uses the branch, year and month filters already on the dashboard. Unit Savings, Monthly Unit Progress and Hours Worked compare the latest two retained weekly snapshots. Agency vs Permanent compares completed months only. Gross Profit is assessed monthly against Outlook and Budget, using `REV - VAR - FIX` for each scenario.

The Unit Count Forecast source is intentionally not part of this flow. Excel Online refresh supports the Power BI-connected queries, but an external workbook link may still require the Outlook calculation to be moved into the OPS workbook or Power Query if Microsoft does not refresh it in the flow. Confirm one scheduled run against the figures in Excel before relying on unattended publishing.
