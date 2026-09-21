type Primitive = string | number | boolean;
type InputRecord = { [key: string]: Primitive | null };
type CellRecord = { [key: string]: Primitive };

type ExportRow = {
    Year: number;
    Month: number;
    Branch: string;
    Group: string;
    Measure: string;
    Value: number;
    Target: number | string;
    Unit: string;
    SnapshotDate: string;
    RecordType: string;
    UpdatedAt: string;
};

type Aggregate = { [key: string]: number };
type ScriptResult = {
    schema: number;
    isTestData: boolean;
    rows: ExportRow[];
    checks: string[];
    warnings: string[];
};

const OPERATIONS_BRANCHES: { [code: string]: string } = {
    AG: "Glasgow",
    LC: "Cambridge",
    NB: "Birmingham",
    NL: "Leeds",
    NM: "Manchester",
    OXF: "Oxford",
    SB: "Bristol",
    SS: "London",
    SW: "Croydon"
};

const FINANCE_BRANCHES: { [code: string]: string } = {
    "GO-EE001": "Cambridge",
    "GO-GL001": "London",
    "GO-GL002": "London",
    "GO-GL003": "Croydon",
    "GO-GL005": "Future Workspace",
    "GO-NE001": "Manchester",
    "GO-NE004": "Leeds",
    "GO-SC005": "Glasgow",
    "GO-SE006": "Oxford",
    "GO-WE008": "Birmingham",
    "GO-WE010": "Bristol"
};

const CONSOLIDATED_ONLY_FINANCE_CODES = new Set([
    "GO-GL006", // UCB Project
    "GO-SE007", // Upper Hayford
    "GO-ZZ102", // Central Overheads
    "GO-ZZ202"  // Central Services
]);

const APPROVED_FINANCE_CODES = new Set([
    ...Object.keys(FINANCE_BRANCHES),
    ...Array.from(CONSOLIDATED_ONLY_FINANCE_CODES)
]);

function main(
    workbook: ExcelScript.Workbook,
    operationsJson: string,
    unitCountJson: string,
    agencyJson: string,
    financeJson: string,
    testMode: boolean
): ScriptResult {
    const refreshedAt = new Date();
    const updatedAt = refreshedAt.toISOString();
    const snapshotDate = updatedAt.slice(0, 10);

    const operations = parseInput(operationsJson, "operationsJson");
    const unitCount = parseInput(unitCountJson, "unitCountJson");
    const agency = parseInput(agencyJson, "agencyJson");
    const finance = parseInput(financeJson, "financeJson");

    const warnings: string[] = [];
    const rows: ExportRow[] = [];

    buildOperationsRows(operations, rows, updatedAt, snapshotDate, warnings);
    const operationsAggregateCount = rows.length;
    const unitAggregates = buildUnitRows(unitCount, rows, updatedAt, snapshotDate, warnings);
    const unitAggregateCount = rows.length - operationsAggregateCount;
    buildAgencyRows(agency, rows, updatedAt, snapshotDate, warnings);
    const agencyAggregateCount = rows.length - operationsAggregateCount - unitAggregateCount;
    const financeStart = rows.length;
    const financeAggregates = buildFinanceRows(finance, rows, updatedAt, snapshotDate, warnings);
    const financeAggregateCount = rows.length - financeStart;
    const unitFinancialStart = rows.length;
    buildMonthlyUnitFinancialRows(unitAggregates, financeAggregates, rows, refreshedAt, updatedAt, snapshotDate);
    const unitFinancialAggregateCount = rows.length - unitFinancialStart;

    if (!rows.length && !testMode) {
        throw new Error("No valid aggregate dashboard rows were produced.");
    }

    // Test mode must never return or retain aggregates derived from live inputs.
    const weeklyRows = testMode ? [] : updateWeeklyHistory(workbook, rows, refreshedAt, updatedAt);
    const checks = [
        "Corrected parser version: 4",
        `Operations query rows: ${operations.length}`,
        `Unit-count query rows: ${unitCount.length}`,
        `Agency query rows: ${agency.length}`,
        `Finance query rows: ${finance.length}`,
        `Operations aggregate rows: ${operationsAggregateCount}`,
        `Unit-count aggregate rows: ${unitAggregateCount}`,
        `Agency aggregate rows: ${agencyAggregateCount}`,
        `Finance aggregate rows: ${financeAggregateCount}`,
        `Monthly Unit Progress financial rows: ${unitFinancialAggregateCount}`,
        `Current aggregate rows created: ${rows.length}`,
        `Weekly history rows returned: ${weeklyRows.length}`,
        "Raw customer/job fields exported to GitHub: 0",
        "Gross Profit formula: Revenue - Variable Costs - Fixed Costs",
        "Monthly Unit Progress Unit Rate uses the simple average of monthly actual unit rates through the prior month",
        "Monthly Unit Progress Target Units uses Outlook revenue less prior-month actual SL-ST006 revenue, divided by the YTD Unit Rate",
        "Actual Units Revenue = Unit Rate x Booked Total",
        "All branches is calculated from the selectable branch outputs only"
    ];

    if (testMode) {
        warnings.unshift("TEST MODE is enabled. Set testMode to false before publishing live figures.");
    }

    return {
        schema: 2,
        isTestData: testMode,
        rows: testMode ? createTestRows(updatedAt, snapshotDate) : rows.concat(weeklyRows),
        checks,
        warnings
    };
}

function parseInput(json: string, parameterName: string): InputRecord[] {
    if (!json || !json.trim()) {
        throw new Error(`${parameterName} was empty.`);
    }
    let parsed: unknown;
    try {
        parsed = JSON.parse(json);
    } catch (error) {
        throw new Error(`${parameterName} was not valid JSON.`);
    }
    if (!Array.isArray(parsed)) {
        throw new Error(`${parameterName} must contain a JSON array.`);
    }
    return (parsed as InputRecord[]).map((record: InputRecord) => normaliseInputRecord(record));
}

function normaliseInputRecord(record: InputRecord): InputRecord {
    const normalised: InputRecord = {};
    for (const key of Object.keys(record)) {
        const trimmed = key.trim();
        const field = trimmed.startsWith("[") && trimmed.endsWith("]")
            ? trimmed.slice(1, -1)
            : trimmed;
        // Store both forms so the script accepts Power BI connector output and
        // manually supplied JSON without depending on one column-name style.
        normalised[field] = record[key];
        normalised[`[${field}]`] = record[key];
    }
    return normalised;
}

function createTestRows(updatedAt: string, snapshotDate: string): ExportRow[] {
    const result: ExportRow[] = [];
    const add = (
        year: number, month: number, group: string, measure: string,
        value: number, unit: string, target: number | string = "",
        recordType: string = "Current", snapshot: string = snapshotDate
    ): void => {
        result.push({
            Year: year, Month: month, Branch: "Demo Branch", Group: group,
            Measure: measure, Value: value, Target: target, Unit: unit,
            SnapshotDate: snapshot, RecordType: recordType, UpdatedAt: updatedAt
        });
    };

    add(2026, 8, "Unit Savings", "Total Unit Savings", 14, "number");
    add(2026, 8, "Monthly Unit Progress", "Booked Labour", 330, "number");
    add(2026, 8, "Monthly Unit Progress", "Booked Vehicle", 110, "number");
    add(2026, 8, "Monthly Unit Progress", "Booked Total", 440, "number");
    add(2026, 8, "Monthly Unit Progress", "Pending Labour", 28, "number");
    add(2026, 8, "Monthly Unit Progress", "Pending Vehicle", 10, "number");
    add(2026, 8, "Monthly Unit Progress", "Pending Total", 38, "number");
    add(2026, 8, "Monthly Unit Progress", "Target Units", 450, "number");
    add(2026, 8, "Monthly Unit Progress", "Unit Rate", 300, "currency");
    add(2026, 8, "Monthly Unit Progress", "Actual Units Revenue", 132000, "currency");
    add(2026, 8, "Hours Worked Breakdown", "Average Hours Worked", 8.8, "number", 9);
    add(2026, 7, "Agency vs Permanent", "Agency %", 0.32, "percent");
    add(2026, 8, "Agency vs Permanent", "Agency %", 0.30, "percent");
    add(2026, 8, "Outlook", "Revenue", 250000, "currency");
    add(2026, 8, "Budget", "Revenue", 265000, "currency");
    add(2026, 8, "Gross Profit", "Actual Revenue", 250000, "currency");
    add(2026, 8, "Gross Profit", "Actual Variable Costs", 95000, "currency");
    add(2026, 8, "Gross Profit", "Actual Fixed Costs", 60000, "currency");
    add(2026, 8, "Gross Profit", "Actual Gross Profit", 95000, "currency");
    add(2026, 8, "Gross Profit", "Actual Gross Margin %", 0.38, "percent");
    add(2026, 8, "Gross Profit", "Outlook Gross Profit", 100000, "currency");
    add(2026, 8, "Gross Profit", "Budget Gross Profit", 105000, "currency");
    add(2026, 8, "Gross Profit", "Variance to Outlook", -5000, "currency");
    add(2026, 8, "Gross Profit", "Variance to Budget", -10000, "currency");

    for (const snapshot of ["2026-08-23", "2026-08-30"]) {
        const later = snapshot === "2026-08-30";
        add(2026, 8, "Unit Savings", "Total Unit Savings", later ? 14 : 12, "number", "", "WeeklySnapshot", snapshot);
        add(2026, 8, "Monthly Unit Progress", "Booked Total", later ? 440 : 420, "number", 450, "WeeklySnapshot", snapshot);
        add(2026, 8, "Hours Worked Breakdown", "Average Hours Worked", later ? 8.8 : 8.6, "number", 9, "WeeklySnapshot", snapshot);
    }
    return result;
}

function numberValue(record: InputRecord, field: string): number {
    const value = record[field];
    if (value === null || value === undefined || value === "") return 0;
    const result = Number(value);
    return Number.isFinite(result) ? result : 0;
}

function textValue(record: InputRecord, field: string): string {
    const value = record[field];
    return value === null || value === undefined ? "" : String(value).trim();
}

function periodFromDate(value: string): { year: number; month: number } | null {
    const match = /^(\d{4})-(\d{2})/.exec(value);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (!Number.isInteger(year) || month < 1 || month > 12) return null;
    return { year, month };
}

function aggregateKey(year: number, month: number, branch: string): string {
    return `${year}|${month}|${branch}`;
}

function getAggregate(map: Map<string, Aggregate>, key: string): Aggregate {
    let value = map.get(key);
    if (!value) {
        value = {};
        map.set(key, value);
    }
    return value;
}

function addValue(record: Aggregate, field: string, value: number): void {
    record[field] = (record[field] || 0) + value;
}

function pushRow(
    rows: ExportRow[], year: number, month: number, branch: string,
    group: string, measure: string, value: number, target: number | string,
    unit: string, updatedAt: string, snapshotDate: string
): void {
    if (!Number.isFinite(value)) return;
    rows.push({
        Year: year,
        Month: month,
        Branch: branch,
        Group: group,
        Measure: measure,
        Value: value,
        Target: target,
        Unit: unit,
        SnapshotDate: snapshotDate,
        RecordType: "Current",
        UpdatedAt: updatedAt
    });
}

function buildOperationsRows(
    source: InputRecord[], rows: ExportRow[], updatedAt: string,
    snapshotDate: string, warnings: string[]
): void {
    const grouped = new Map<string, Aggregate>();
    const unknownCodes = new Set<string>();

    for (const record of source) {
        const suppliedYear = numberValue(record, "[Year]");
        const suppliedMonth = numberValue(record, "[Month]");
        const period = Number.isInteger(suppliedYear) && suppliedMonth >= 1 && suppliedMonth <= 12
            ? { year: suppliedYear, month: suppliedMonth }
            : periodFromDate(textValue(record, "[Date]"));
        const code = textValue(record, "[BranchCode]");
        const branch = OPERATIONS_BRANCHES[code];
        if (!period || !branch) {
            if (code && !branch) unknownCodes.add(code);
            continue;
        }
        const aggregate = getAggregate(grouped, aggregateKey(period.year, period.month, branch));
        addValue(aggregate, "labourActual", numberValue(record, "[LabourActual]"));
        addValue(aggregate, "vehicleActual", numberValue(record, "[VehicleActual]"));
        addValue(aggregate, "labourSaving", numberValue(record, "[LabourSaving]"));
        addValue(aggregate, "vehicleSaving", numberValue(record, "[VehicleSaving]"));
        addValue(aggregate, "estimatedHours", numberValue(record, "[EstimatedHours]"));
        addValue(aggregate, "actualHours", numberValue(record, "[ActualHours]"));
        addValue(aggregate, "paidHours", numberValue(record, "[PaidHours]"));
    }

    grouped.forEach((a: Aggregate, key: string) => {
        const [yearText, monthText, branch] = key.split("|");
        const year = Number(yearText), month = Number(monthText);
        const totalActual = a.labourActual + a.vehicleActual;
        const totalSaving = a.labourSaving + a.vehicleSaving;
        const underUtilised = a.paidHours - a.actualHours;
        const averageHours = a.labourActual === 0 ? 0 : a.actualHours / a.labourActual;

        pushRow(rows, year, month, branch, "Unit Savings", "Labour Units Actual", a.labourActual, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Vehicle Units Actual", a.vehicleActual, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Labour Unit Savings", a.labourSaving, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Vehicle Unit Savings", a.vehicleSaving, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Total Unit Savings", totalSaving, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Labour Savings %", a.labourActual === 0 ? 0 : a.labourSaving / a.labourActual, "", "percent", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Vehicle Savings %", a.vehicleActual === 0 ? 0 : a.vehicleSaving / a.vehicleActual, "", "percent", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Unit Savings", "Total Savings %", totalActual === 0 ? 0 : totalSaving / totalActual, "", "percent", updatedAt, snapshotDate);

        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Estimated Hours", a.estimatedHours, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Actual Hours", a.actualHours, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Hours Paid", a.paidHours, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Hours Under Utilised", underUtilised, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Under Utilised %", a.paidHours === 0 ? 0 : underUtilised / a.paidHours, "", "percent", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Hours Worked Breakdown", "Average Hours Worked", averageHours, 9, "number", updatedAt, snapshotDate);
    });

    for (const code of Array.from(unknownCodes).sort()) {
        warnings.push(`New operations branch code detected: ${code}. Review its branch mapping.`);
    }
}

function buildUnitRows(
    source: InputRecord[], rows: ExportRow[], updatedAt: string,
    snapshotDate: string, warnings: string[]
): Map<string, Aggregate> {
    const grouped = new Map<string, Aggregate>();
    const unknownCodes = new Set<string>();
    const unknownStatuses = new Set<string>();

    for (const record of source) {
        const suppliedYear = numberValue(record, "[Year]");
        const suppliedMonth = numberValue(record, "[Month]");
        const period = Number.isInteger(suppliedYear) && suppliedMonth >= 1 && suppliedMonth <= 12
            ? { year: suppliedYear, month: suppliedMonth }
            : periodFromDate(textValue(record, "[Date]"));
        const code = textValue(record, "[BranchCode]");
        const branch = OPERATIONS_BRANCHES[code];
        const unitType = textValue(record, "[UnitType]");
        const status = textValue(record, "[Status]").toUpperCase();
        const quantity = numberValue(record, "[Quantity]");
        const hasCalculatedCategories = record["[BookedQuantity]"] !== undefined || record["[PendingQuantity]"] !== undefined;
        if (!period || !branch) {
            if (code && !branch) unknownCodes.add(code);
            continue;
        }
        if (!hasCalculatedCategories && !["C", "L", "P", "W"].includes(status)) {
            if (status) unknownStatuses.add(status);
            continue;
        }
        if (unitType !== "Labour" && unitType !== "Vehicle") continue;
        const aggregate = getAggregate(grouped, aggregateKey(period.year, period.month, branch));
        if (hasCalculatedCategories) {
            addValue(aggregate, unitType === "Labour" ? "bookedLabour" : "bookedVehicle", numberValue(record, "[BookedQuantity]"));
            addValue(aggregate, unitType === "Labour" ? "pendingLabour" : "pendingVehicle", numberValue(record, "[PendingQuantity]"));
        } else {
            if (status === "W") addValue(aggregate, unitType === "Labour" ? "bookedLabour" : "bookedVehicle", quantity);
            if (status === "P") addValue(aggregate, unitType === "Labour" ? "pendingLabour" : "pendingVehicle", quantity);
        }
    }

    grouped.forEach((a: Aggregate, key: string) => {
        const [yearText, monthText, branch] = key.split("|");
        const year = Number(yearText), month = Number(monthText);
        const bookedLabour = a.bookedLabour || 0;
        const bookedVehicle = a.bookedVehicle || 0;
        const pendingLabour = a.pendingLabour || 0;
        const pendingVehicle = a.pendingVehicle || 0;
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Booked Labour", bookedLabour, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Booked Vehicle", bookedVehicle, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Booked Total", bookedLabour + bookedVehicle, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Pending Labour", pendingLabour, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Pending Vehicle", pendingVehicle, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Pending Total", pendingLabour + pendingVehicle, "", "number", updatedAt, snapshotDate);
    });

    for (const code of Array.from(unknownCodes).sort()) {
        warnings.push(`New unit-count branch code detected: ${code}. Review its branch mapping.`);
    }
    for (const status of Array.from(unknownStatuses).sort()) {
        warnings.push(`New unit status detected: ${status}. Review whether it is booked, pending or excluded.`);
    }
    return grouped;
}

function agencyBranch(sourceBranch: string): string {
    if (sourceBranch === "Thurrock" || sourceBranch === "Silvertown" || sourceBranch === "Logistic City") return "London";
    return sourceBranch;
}

function buildAgencyRows(
    source: InputRecord[], rows: ExportRow[], updatedAt: string,
    snapshotDate: string, warnings: string[]
): void {
    const grouped = new Map<string, Aggregate>();
    const statuses = new Set(["Agency", "Permanent"]);
    const unknownStatuses = new Set<string>();
    for (const record of source) {
        const period = periodFromDate(textValue(record, "[Date]"));
        const branch = agencyBranch(textValue(record, "[Branch]"));
        const status = textValue(record, "[Status]");
        if (!period || !branch) continue;
        if (!statuses.has(status)) {
            if (status) unknownStatuses.add(status);
            continue;
        }
        const aggregate = getAggregate(grouped, aggregateKey(period.year, period.month, branch));
        addValue(aggregate, status === "Agency" ? "agencyHours" : "permanentHours", numberValue(record, "[Hours]"));
    }
    grouped.forEach((a: Aggregate, key: string) => {
        const [yearText, monthText, branch] = key.split("|");
        const year = Number(yearText), month = Number(monthText);
        const agencyHours = a.agencyHours || 0;
        const permanentHours = a.permanentHours || 0;
        const total = agencyHours + permanentHours;
        pushRow(rows, year, month, branch, "Agency vs Permanent", "Agency Hours", agencyHours, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Agency vs Permanent", "Permanent Hours", permanentHours, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Agency vs Permanent", "Agency %", total === 0 ? 0 : agencyHours / total, "", "percent", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Agency vs Permanent", "Permanent %", total === 0 ? 0 : permanentHours / total, "", "percent", updatedAt, snapshotDate);
    });
    for (const status of Array.from(unknownStatuses).sort()) {
        warnings.push(`New agency status detected: ${status}. Review its treatment.`);
    }
}

function addFinanceValues(a: Aggregate, record: InputRecord): void {
    const actualRevenue = -numberValue(record, "[Actual]");
    const outlookRevenue = -numberValue(record, "[Forecast]");
    const budgetRevenue = -numberValue(record, "[Budget]");
    addValue(a, "actualRevenue", actualRevenue);
    addValue(a, "outlookRevenue", outlookRevenue);
    addValue(a, "budgetRevenue", budgetRevenue);
    addValue(a, "actualVariable", numberValue(record, "[ActualVariableCost]"));
    addValue(a, "outlookVariable", numberValue(record, "[ForecastVariableCost]"));
    addValue(a, "budgetVariable", numberValue(record, "[BudgetVariableCost]"));
    addValue(a, "actualFixed", numberValue(record, "[ActualFixedCost]"));
    addValue(a, "outlookFixed", numberValue(record, "[ForecastFixedCost]"));
    addValue(a, "budgetFixed", numberValue(record, "[BudgetFixedCost]"));

    const revenueType = textValue(record, "[RevenueType]");
    // The workbook's Unit Rate uses actual revenue excluding SL-ST006.
    if (revenueType !== "SL-ST006") {
        addValue(a, "unitRateActualRevenue", actualRevenue);
    } else {
        // The Monthly Unit Progress target removes the prior month's
        // actual SL-ST006 revenue from the selected month's Outlook revenue.
        addValue(a, "storageActualRevenue", actualRevenue);
    }
}

function buildFinanceRows(
    source: InputRecord[], rows: ExportRow[], updatedAt: string,
    snapshotDate: string, warnings: string[]
): Map<string, Aggregate> {
    const grouped = new Map<string, Aggregate>();
    const unknownCodes = new Set<string>();
    let blankGeographyRows = 0;
    for (const record of source) {
        const year = numberValue(record, "[Year]");
        const month = numberValue(record, "[Month]");
        const code = textValue(record, "[GeographyCode]");
        if (!Number.isInteger(year) || month < 1 || month > 12) continue;
        if (!code) {
            blankGeographyRows++;
            continue;
        }
        if (!APPROVED_FINANCE_CODES.has(code)) unknownCodes.add(code);


        const branch = FINANCE_BRANCHES[code];
        if (branch) addFinanceValues(getAggregate(grouped, aggregateKey(year, month, branch)), record);
    }

    grouped.forEach((a: Aggregate, key: string) => {
        const [yearText, monthText, branch] = key.split("|");
        const year = Number(yearText), month = Number(monthText);
        const actualGp = a.actualRevenue - a.actualVariable - a.actualFixed;
        const outlookGp = a.outlookRevenue - a.outlookVariable - a.outlookFixed;
        const budgetGp = a.budgetRevenue - a.budgetVariable - a.budgetFixed;
        pushRow(rows, year, month, branch, "Outlook", "Revenue", a.outlookRevenue, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Budget", "Revenue", a.budgetRevenue, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Actual Revenue", a.actualRevenue, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Actual Variable Costs", a.actualVariable, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Actual Fixed Costs", a.actualFixed, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Actual Gross Profit", actualGp, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Actual Gross Margin %", a.actualRevenue === 0 ? 0 : actualGp / a.actualRevenue, "", "percent", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Outlook Revenue", a.outlookRevenue, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Outlook Variable Costs", a.outlookVariable, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Outlook Fixed Costs", a.outlookFixed, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Outlook Gross Profit", outlookGp, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Budget Revenue", a.budgetRevenue, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Budget Variable Costs", a.budgetVariable, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Budget Fixed Costs", a.budgetFixed, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Budget Gross Profit", budgetGp, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Variance to Outlook", actualGp - outlookGp, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Gross Profit", "Variance to Budget", actualGp - budgetGp, "", "currency", updatedAt, snapshotDate);
    });

    for (const code of Array.from(unknownCodes).sort()) {
        warnings.push(`NEW GEOGRAPHY CODE: ${code}. Excluded from the selectable branch totals; review whether it should become selectable.`);
    }
    if (blankGeographyRows) {
        warnings.push(`Finance rows with blank GeographyCode: ${blankGeographyRows}. Excluded pending review.`);
    }
    return grouped;
}

function buildMonthlyUnitFinancialRows(
    unitAggregates: Map<string, Aggregate>,
    financeAggregates: Map<string, Aggregate>,
    rows: ExportRow[],
    refreshedAt: Date,
    updatedAt: string,
    snapshotDate: string
): void {
    const currentYear = refreshedAt.getUTCFullYear();
    const currentMonth = refreshedAt.getUTCMonth() + 1;
    const branchYears = new Set<string>();

    for (const key of Array.from(unitAggregates.keys())) {
        const [yearText, , branch] = key.split("|");
        branchYears.add(`${yearText}|${branch}`);
    }

    const unitRates = new Map<string, number>();
    for (const branchYear of Array.from(branchYears)) {
        const splitAt = branchYear.indexOf("|");
        const year = Number(branchYear.slice(0, splitAt));
        const branch = branchYear.slice(splitAt + 1);
        const endMonth = year < currentYear ? 12 : year === currentYear ? Math.max(0, currentMonth - 1) : 0;
        const monthlyRates: number[] = [];

        for (let month = 1; month <= endMonth; month++) {
            const key = aggregateKey(year, month, branch);
            const units = unitAggregates.get(key);
            const finance = financeAggregates.get(key);
            if (!units || !finance) continue;
            const bookedTotal = (units.bookedLabour || 0) + (units.bookedVehicle || 0);
            if (bookedTotal <= 0) continue;
            const monthlyRate = (finance.unitRateActualRevenue || 0) / bookedTotal;
            if (Number.isFinite(monthlyRate)) monthlyRates.push(monthlyRate);
        }

        if (monthlyRates.length) {
            unitRates.set(branchYear, monthlyRates.reduce((sum, value) => sum + value, 0) / monthlyRates.length);
        }
    }

    for (const [key, units] of Array.from(unitAggregates.entries())) {
        const [yearText, monthText, branch] = key.split("|");
        const year = Number(yearText), month = Number(monthText);
        const rateYear = year > currentYear ? currentYear : year;
        const unitRate = unitRates.get(`${rateYear}|${branch}`);
        if (unitRate === undefined || !Number.isFinite(unitRate) || unitRate === 0) continue;

        const bookedTotal = (units.bookedLabour || 0) + (units.bookedVehicle || 0);
        const finance = financeAggregates.get(key);
        const outlookRevenue = finance ? (finance.outlookRevenue || 0) : 0;
        const priorKey = month > 1
            ? aggregateKey(year, month - 1, branch)
            : aggregateKey(year - 1, 12, branch);
        const priorFinance = financeAggregates.get(priorKey);
        const priorStorageActual = priorFinance ? (priorFinance.storageActualRevenue || 0) : 0;
        const targetRevenue = outlookRevenue - priorStorageActual;
        const targetUnits = targetRevenue / unitRate;
        const actualUnitsRevenue = unitRate * bookedTotal;

        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Target Units", targetUnits, "", "number", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Unit Rate", unitRate, "", "currency", updatedAt, snapshotDate);
        pushRow(rows, year, month, branch, "Monthly Unit Progress", "Actual Units Revenue", actualUnitsRevenue, "", "currency", updatedAt, snapshotDate);
    }
}

function recordFromValues(headers: string[], values: Primitive[]): CellRecord {
    const result: CellRecord = {};
    for (let i = 0; i < headers.length; i++) result[headers[i]] = values[i] === undefined ? "" : values[i];
    return result;
}

function historyKey(record: CellRecord): string {
    return [record.SnapshotDate, record.Year, record.Month, record.Branch, record.Group, record.Measure].join("|");
}

function updateWeeklyHistory(
    workbook: ExcelScript.Workbook,
    currentRows: ExportRow[],
    refreshedAt: Date,
    updatedAt: string
): ExportRow[] {
    const weeklyGroups = new Set(["Unit Savings", "Monthly Unit Progress", "Hours Worked Breakdown"]);
    const eligible = currentRows.filter(r => weeklyGroups.has(r.Group));
    if (!eligible.length) return [];
    const latestPeriod = Math.max(...eligible.map(r => r.Year * 100 + r.Month));
    const weekEndingDate = new Date(Date.UTC(refreshedAt.getUTCFullYear(), refreshedAt.getUTCMonth(), refreshedAt.getUTCDate()));
    const daysToSunday = (7 - weekEndingDate.getUTCDay()) % 7;
    weekEndingDate.setUTCDate(weekEndingDate.getUTCDate() + daysToSunday);
    const weekEnding = weekEndingDate.toISOString().slice(0, 10);
    const snapshot: ExportRow[] = eligible
        .filter(r => r.Year * 100 + r.Month === latestPeriod)
        .map(r => ({ ...r, SnapshotDate: weekEnding, RecordType: "WeeklySnapshot" }));

    let history = workbook.getTable("OpsDashboardHistory");
    if (!history) {
        const sheet = workbook.getWorksheet("Dashboard History") || workbook.addWorksheet("Dashboard History");
        sheet.getRange("A1:K1").setValues([[
            "Year", "Month", "Branch", "Group", "Measure", "Value",
            "Target", "Unit", "SnapshotDate", "RecordType", "UpdatedAt"
        ]]);
        history = sheet.addTable("A1:K1", true);
        history.setName("OpsDashboardHistory");
    }

    const headers = history.getHeaderRowRange().getTexts()[0];
    const body = history.getRowCount() ? history.getRangeBetweenHeaderAndTotal() : null;
    const existingValues: Primitive[][] = body ? body.getValues() as Primitive[][] : [];
    const existing: CellRecord[] = existingValues.map(values => recordFromValues(headers, values));
    const indexByKey = new Map<string, number>();
    existing.forEach((record, index) => indexByKey.set(historyKey(record), index));
    const additions: Primitive[][] = [];
    let historyChanged = false;

    for (const row of snapshot) {
        const record = row as unknown as CellRecord;
        const values = headers.map(header => record[header] === undefined ? "" : record[header]) as Primitive[];
        const index = indexByKey.get(historyKey(record));
        if (index === undefined) {
            additions.push(values);
        } else {
            existingValues[index] = values;
            historyChanged = true;
        }
    }

    // Write existing weekly snapshot updates in one operation instead of one
    // setValues call per row. This keeps RunScript comfortably below the
    // Power Automate/Office Scripts HTTP timeout on larger weekly snapshots.
    if (body && historyChanged) body.setValues(existingValues);
    if (additions.length) history.addRows(-1, additions);

    const merged = new Map<string, CellRecord>();
    for (const record of existing) merged.set(historyKey(record), record);
    for (const row of snapshot) {
        const record = row as unknown as CellRecord;
        merged.set(historyKey(record), record);
    }
    const cutoff = new Date(refreshedAt);
    cutoff.setUTCFullYear(cutoff.getUTCFullYear() - 2);
    const retained: ExportRow[] = [];
    for (const record of Array.from(merged.values())) {
        if (String(record.RecordType) !== "WeeklySnapshot") continue;
        const dateText = String(record.SnapshotDate).slice(0, 10);
        if (new Date(`${dateText}T00:00:00Z`) < cutoff) continue;
        retained.push({
            Year: Number(record.Year),
            Month: Number(record.Month),
            Branch: String(record.Branch),
            Group: String(record.Group),
            Measure: String(record.Measure),
            Value: Number(record.Value),
            Target: record.Target === "" ? "" : Number(record.Target),
            Unit: String(record.Unit || "number"),
            SnapshotDate: dateText,
            RecordType: "WeeklySnapshot",
            UpdatedAt: String(record.UpdatedAt || updatedAt)
        });
    }
    return retained;
}