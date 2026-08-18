/* global Excel */

import { FilterPlan } from "./parser";

export interface TableSnapshot {
  headers: any[];
  rows: any[][];
}

// Grabs the first Table on the active sheet, auto-creating one from the
// sheet's data if it isn't already formatted as a Table. This is what lets
// someone paste or upload *any* wine spreadsheet and start filtering right
// away instead of first having to select their data and press Ctrl+T.
//
// This add-in assumes one wine table per sheet — if you have multiple
// tables, change getItemAt(0) to tables.getItem("YourTableName").
async function getFirstTable(context: Excel.RequestContext): Promise<Excel.Table> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const tables = sheet.tables;
  tables.load("items");
  await context.sync();

  if (tables.items.length > 0) {
    return tables.items[0];
  }

  return createTableFromUsedRange(context, sheet);
}

// Finds the sheet's used range, skips any leading blank rows (e.g. a title
// row above the data), and turns the remaining block into an Excel Table
// with the first surviving row treated as headers.
async function createTableFromUsedRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet
): Promise<Excel.Table> {
  const usedRange = sheet.getUsedRangeOrNullObject(true);
  usedRange.load("values, rowCount, columnCount, isNullObject");
  await context.sync();

  if (usedRange.isNullObject || usedRange.rowCount < 2) {
    throw new Error(
      "Couldn't find any wine data on this sheet. Add your list (with a header row and at least one wine underneath) and try again."
    );
  }

  const values = usedRange.values as any[][];
  const isBlankRow = (row: any[]) => row.every((c) => c === "" || c === null || c === undefined);

  let headerRowOffset = values.findIndex((row) => !isBlankRow(row));
  if (headerRowOffset === -1) headerRowOffset = 0;

  if (headerRowOffset >= values.length - 1) {
    throw new Error(
      "Couldn't find a header row with wine data underneath it. Check your sheet and try again."
    );
  }

  const remainingRows = values.length - headerRowOffset;
  const dataRange =
    headerRowOffset === 0
      ? usedRange
      : usedRange.getCell(headerRowOffset, 0).getResizedRange(remainingRows - 1, usedRange.columnCount - 1);

  const table = sheet.tables.add(dataRange, true);
  table.name = "WineList";

  // Excel applies its own default banded-blue style to any newly created
  // Table — that's Excel's behavior on tables.add(), not something we asked
  // for. Since we auto-created this table on the user's behalf (they didn't
  // press Ctrl+T themselves), reset it to "no style" so filtering/sorting
  // works without silently reformatting their sheet. A table the user
  // already made themselves is left completely alone (see getFirstTable).
  table.style = "TableStyleNone";

  table.load("items");
  await context.sync();

  return table;
}

export async function getTableSnapshot(): Promise<TableSnapshot> {
  return Excel.run(async (context) => {
    const table = await getFirstTable(context);
    const headerRange = table.getHeaderRowRange();
    const bodyRange = table.getDataBodyRange();
    headerRange.load("values");
    bodyRange.load("values");
    await context.sync();

    return {
      headers: headerRange.values[0],
      rows: bodyRange.values
    };
  });
}

// Applies filters + sort via Excel's native Table API. This never rewrites a
// single cell's value or formatting — it only shows/hides and reorders rows,
// exactly like clicking the header dropdown arrows yourself. Returns the
// number of visible data rows after the filter is applied.
export async function applyFilterPlan(plan: FilterPlan): Promise<number> {
  return Excel.run(async (context) => {
    const table = await getFirstTable(context);

    // Start from a clean slate so re-applying a new prompt doesn't stack
    // on top of a previous filter.
    table.clearFilters();
    await context.sync();

    plan.valuesFilters.forEach((vf) => {
      if (!vf.values.length) return;
      const column = table.columns.getItem(vf.column);
      column.filter.applyValuesFilter(vf.values);
    });

    plan.customFilters.forEach((cf) => {
      const column = table.columns.getItem(cf.column);
      column.filter.applyCustomFilter(cf.criteria1, cf.criteria2, cf.oper);
    });

    await context.sync();

    if (plan.sort) {
      const sortColumn = table.columns.getItem(plan.sort.column);
      sortColumn.load("index");
      await context.sync();

      table.sort.apply([{ key: sortColumn.index, ascending: plan.sort.ascending }]);
      await context.sync();
    }

    const visible = table.getDataBodyRange().getVisibleView();
    visible.load("rowCount");
    await context.sync();

    return visible.rowCount;
  });
}

// "Clear Filters" — native Table.clearFilters(), so every row reappears with
// its original formatting completely untouched (nothing was ever rewritten).
export async function clearAllFilters(): Promise<void> {
  return Excel.run(async (context) => {
    const table = await getFirstTable(context);
    table.clearFilters();
    await context.sync();
  });
}