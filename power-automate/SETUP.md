# Weekly analysis refresh setup

This design keeps the calculation workbook in SharePoint/OneDrive and publishes only the `Dashboard Export` aggregate table.

1. Store the OPS workbook and `HG Outlook - Mar 26` in SharePoint/OneDrive locations accessible to the flow owner. Keep the Outlook link stable.
2. Add the Office Script in `office-scripts/refresh-and-export.ts` to Excel for the web.
3. Create a scheduled cloud flow with a Recurrence trigger for once a week, after the weekly OPS figures are final, in the `Europe/London` time zone. Friday evening is recommended if the source data is complete by then.
4. Add **Excel Online (Business) → Run script**, select the OPS workbook, and run `refresh-and-export`.
5. Add a GitHub **Repository dispatch** call for `HarrowGreenltd/Operations-KPI-Dashboard` with event type `ops-kpi-refresh`. Send the script result as the `client_payload` body.
6. Test with `isTestData: true` while the site is public. Do not publish real OPS figures until private GitHub Pages is available and access is configured.
7. Enable the flow only after the workbook refresh and dashboard checks both pass. The first successful run creates the baseline; prior-week commentary becomes available after the second weekly snapshot.

The Office Script creates and maintains an `OpsDashboardHistory` table in a `Dashboard History` worksheet. It retains aggregate weekly snapshots for Unit Savings, Monthly Unit Progress and Hours Worked; no raw job or customer data is added. This avoids granting the GitHub workflow permission to write back to the repository.

The Analysis tab uses the branch, year and month filters already on the dashboard. Unit Savings, Monthly Unit Progress and Hours Worked compare the latest two retained weekly snapshots. Agency vs Permanent compares completed months only. Gross Profit is assessed monthly against Outlook and Budget, using `REV - VAR - FIX` for each scenario.

The Unit Count Forecast source is intentionally not part of this flow. Excel Online refresh supports the Power BI-connected queries, but an external workbook link may still require the Outlook calculation to be moved into the OPS workbook or Power Query if Microsoft does not refresh it in the flow. Confirm one scheduled run against the figures in Excel before relying on unattended publishing.
