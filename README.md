# @mshindi-labs/gsheets

A fully-typed TypeScript wrapper around the Google Sheets v4 API. Every method returns a discriminated-union response — no thrown exceptions, no manual error parsing.

```ts
const result = await sheet.getHeaderRow('Customers');

if (result.success) {
  console.log(result.data); // string[]
} else {
  console.error(result.problem); // string
}
```

---

## Table of Contents

- [Installation](#installation)
- [Setup](#setup)
  - [Step 1 — Google Cloud credentials](#step-1--google-cloud-credentials)
  - [Step 2 — Obtain a refresh token](#step-2--obtain-a-refresh-token)
  - [Step 3 — Create a client](#step-3--create-a-client)
- [Response shape](#response-shape)
- [Methods](#methods)
  - [Metadata & structure](#metadata--structure)
  - [Reading data](#reading-data)
  - [Writing data](#writing-data)
  - [Batch operations](#batch-operations)
  - [Rows & columns](#rows--columns)
  - [Sheet management](#sheet-management)
  - [Formatting & styling](#formatting--styling)
  - [Data utilities](#data-utilities)
  - [Filtering & sorting](#filtering--sorting)
  - [Named ranges](#named-ranges)
- [TypeScript types](#typescript-types)
- [Requirements](#requirements)
- [License](#license)

---

## Installation

```bash
npm install @mshindi-labs/gsheets
# or
pnpm add @mshindi-labs/gsheets
# or
yarn add @mshindi-labs/gsheets
```

---

## Setup

### Step 1 — Google Cloud credentials

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project (or select an existing one).
3. Enable the **Google Sheets API** and **Google Drive API**.
4. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**.
5. Choose **Desktop app** (for local/server scripts) or **Web application**.
6. Download the JSON — you need `client_id`, `client_secret`, and your chosen `redirect_uri`.

### Step 2 — Obtain a refresh token

Run this once to exchange an authorization code for a long-lived refresh token.

```ts
import { createOAuth2Client, generateAuthUrl, retrieveRefreshToken } from '@mshindi-labs/gsheets';

const client = createOAuth2Client({
  clientId: process.env.GOOGLE_CLIENT_ID!,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
  redirectUri: 'http://localhost',
});

// 1. Print the URL, open it in a browser, grant access.
console.log(generateAuthUrl(client));

// 2. Paste the code= query param from the redirect URL here.
const refreshToken = await retrieveRefreshToken(client, 'PASTE_CODE_HERE');
console.log('Save this refresh token:', refreshToken);
```

Store the refresh token securely (e.g. in an environment variable or secrets manager). It does not expire unless access is explicitly revoked.

### Step 3 — Create a client

```ts
import { createSheetsClient } from '@mshindi-labs/gsheets';

const sheet = createSheetsClient(
  {
    clientId:     process.env.GOOGLE_CLIENT_ID!,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET!,
    redirectUri:  process.env.GOOGLE_REDIRECT_URI!,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN!,
  },
  'YOUR_SPREADSHEET_ID', // from the spreadsheet URL: /spreadsheets/d/<ID>/edit
);
```

`createSheetsClient` is a convenience factory. If you already manage your own OAuth2 client you can instantiate the class directly:

```ts
import { GoogleSpreadSheet } from '@mshindi-labs/gsheets';
import { google } from 'googleapis';

const auth = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
auth.setCredentials({ refresh_token: refreshToken });

const sheet = new GoogleSpreadSheet({ auth, spreadsheetId: 'YOUR_ID' });
```

---

## Response shape

Every method returns `Promise<GoogleSheetResponse<T>>`, a discriminated union:

```ts
type GoogleSheetResponse<T> =
  | { success: true;  data: T }
  | { success: false; problem: string };
```

Check `result.success` before accessing `result.data`. The `problem` string on failures includes both a human-readable context label and the underlying API error message.

---

## Methods

### Metadata & structure

#### `getMetadata()`

Returns the spreadsheet's top-level properties (title, locale, time zone) along with the properties of every sheet it contains.

```ts
const res = await sheet.getMetadata();
if (res.success) console.log(res.data.properties?.title);
```

---

#### `getTabs()`

Returns a lightweight summary of every tab: title, numeric `sheetId`, index, and grid dimensions. This is the cheapest way to discover available sheets.

```ts
const res = await sheet.getTabs();
if (res.success) {
  res.data.forEach(tab => console.log(tab.title, tab.sheetId));
}
```

---

#### `getSheetByName(sheetTitle)`

Finds a sheet by its tab title and returns the full `Sheet` resource.

```ts
const res = await sheet.getSheetByName('Inventory');
if (res.success) console.log(res.data.properties?.sheetId);
```

---

### Reading data

#### `getDataRange(props)`

Reads a rectangular range and returns the full `ValueRange` resource.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `sheetTitle` | `string` | — | Tab title (use this or `sheetId`) |
| `sheetId` | `number` | — | Numeric sheet ID |
| `range` | `string` | — | A1 notation, e.g. `'A1:D20'` |
| `valueRenderOption` | `string` | `'FORMATTED_VALUE'` | How cell values are rendered |

```ts
const res = await sheet.getDataRange({ sheetTitle: 'Orders', range: 'A1:E100' });
if (res.success) console.log(res.data.values);
```

---

#### `getLastRowWithData(sheetTitle, column?)`

Returns the 1-based row number of the last row that contains data in the given column (defaults to column A).

```ts
const res = await sheet.getLastRowWithData('Sheet1', 'C');
if (res.success) console.log('Last row:', res.data); // e.g. 42
```

---

#### `getRow({ sheetTitle, rowNumber, valueRenderOption? })`

Reads all values from a single row and returns them as a flat string array. Row numbers are 1-based.

```ts
const res = await sheet.getRow({ sheetTitle: 'Sales', rowNumber: 5 });
if (res.success) console.log(res.data); // ['Alice', '2024-01-15', '1200']
```

---

#### `getValuesBySheetTitle(sheetTitle, range, valueRenderOption?)`

Reads a range and returns the raw 2-D values array.

```ts
const res = await sheet.getValuesBySheetTitle('Config', 'A1:B20');
```

---

#### `getColumnValues(sheetTitle, column, valueRenderOption?)`

Reads an entire column and returns a 2-D array where each element is `[cellValue]`.

```ts
const res = await sheet.getColumnValues('Employees', 'B');
if (res.success) res.data.forEach(([name]: string[]) => console.log(name));
```

---

### Writing data

#### `appendData(props)`

Appends rows to the end of the sheet's existing dataset.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `sheetTitle` | `string` | — | Target tab |
| `data` | `string[][]` | — | Rows to append |
| `valueInputOption` | `string` | `'USER_ENTERED'` | `'USER_ENTERED'` parses formulas/dates; `'RAW'` stores as-is |
| `insertDataOption` | `string` | `'INSERT_ROWS'` | `'INSERT_ROWS'` shifts existing rows; `'OVERWRITE'` fills blank cells |

```ts
await sheet.appendData({
  sheetTitle: 'Log',
  data: [['2024-06-01', 'Deploy', 'success']],
});
```

---

#### `updateRange(props)`

Overwrites a range with new values.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `sheetTitle` | `string` | — | Target tab |
| `range` | `string` | — | A1 notation start cell or range |
| `data` | `string[][]` | — | Values to write |
| `valueInputOption` | `string` | `'USER_ENTERED'` | — |

```ts
await sheet.updateRange({
  sheetTitle: 'Summary',
  range: 'B2',
  data: [['Updated value']],
});
```

---

#### `clearRange(sheetTitle, range)`

Erases all values in a range while leaving formatting intact.

```ts
await sheet.clearRange('Temp', 'A1:Z1000');
```

---

### Batch operations

Batch methods reduce quota consumption and round-trip latency by bundling multiple operations into a single API call.

#### `batchGetRanges(props)`

Fetches multiple ranges in one call.

```ts
const res = await sheet.batchGetRanges({
  sheetTitle: 'Dashboard',
  ranges: ['A1:A10', 'C1:C10', 'E1:E10'],
});
if (res.success) res.data.valueRanges?.forEach(vr => console.log(vr.values));
```

---

#### `batchUpdateRanges(props)`

Writes to multiple ranges in one call. Each entry has its own range and data matrix.

```ts
await sheet.batchUpdateRanges({
  sheetTitle: 'Config',
  entries: [
    { range: 'B2', data: [['v1.2.0']] },
    { range: 'B3', data: [['2024-06-01']] },
  ],
});
```

---

#### `batchClearRanges(props)`

Clears values from multiple ranges in one call. Formatting is preserved.

```ts
await sheet.batchClearRanges({
  sheetTitle: 'Cache',
  ranges: ['A2:A100', 'D2:D100'],
});
```

---

### Rows & columns

All indices are **0-based** and ranges are half-open `[start, end)` matching the Sheets API convention.

#### `insertRows(sheetTitle, startIndex, endIndex)`

Inserts blank rows, shifting existing rows down.

```ts
// Insert 3 blank rows before the current row 2
await sheet.insertRows('Sheet1', 1, 4);
```

---

#### `insertColumns(sheetTitle, startIndex, endIndex)`

Inserts blank columns, shifting existing columns right.

```ts
// Insert 1 blank column before column B
await sheet.insertColumns('Sheet1', 1, 2);
```

---

#### `deleteRows(sheetTitle, startIndex, endIndex)`

Permanently deletes rows. **Irreversible.**

```ts
// Delete rows 2–4 (0-based: start=1, end=4)
await sheet.deleteRows('Sheet1', 1, 4);
```

---

#### `deleteColumns(sheetTitle, startIndex, endIndex)`

Permanently deletes columns. **Irreversible.**

```ts
// Delete columns C and D (0-based: start=2, end=4)
await sheet.deleteColumns('Sheet1', 2, 4);
```

---

### Sheet management

#### `addSheet(props)`

Creates a new empty sheet tab.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `title` | `string` | — | Must be unique within the spreadsheet |
| `index` | `number` | end | 0-based tab position |
| `rowCount` | `number` | 1000 | Initial row count |
| `columnCount` | `number` | 26 | Initial column count |

```ts
await sheet.addSheet({ title: 'Q3 Report', rowCount: 500 });
```

---

#### `deleteSheet(sheetTitle)`

Permanently deletes a sheet and all its data. **Irreversible.**

```ts
await sheet.deleteSheet('Old Data');
```

---

#### `duplicateSheet(sheetId, newTitle?)`

Creates a full copy of an existing sheet including data, formatting, and formulas.

```ts
const tabs = await sheet.getTabs();
const templateId = tabs.data?.find(t => t.title === 'Template')?.sheetId;
await sheet.duplicateSheet(templateId!, 'January Report');
```

---

#### `renameSheet(currentTitle, newTitle)`

Renames a sheet tab.

```ts
await sheet.renameSheet('Sheet1', 'Customers');
```

---

#### `hideSheet(sheetTitle)` / `showSheet(sheetTitle)`

Hides or shows a sheet tab. Data is preserved when hidden.

```ts
await sheet.hideSheet('Internal Notes');
await sheet.showSheet('Internal Notes');
```

---

#### `moveSheet(sheetTitle, newIndex)`

Moves a sheet to a new position in the tab bar (0 = leftmost).

```ts
await sheet.moveSheet('Summary', 0);
```

---

#### `protectSheet(props)`

Adds a protection rule to an entire sheet.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `sheetTitle` | `string` | — | Sheet to protect |
| `description` | `string` | — | Optional note explaining the protection |
| `warningOnly` | `boolean` | `false` | Show a warning instead of blocking edits |
| `editorEmails` | `string[]` | — | Emails allowed to edit despite protection |

```ts
await sheet.protectSheet({
  sheetTitle: 'Master Config',
  description: 'Do not edit without approval',
  editorEmails: ['admin@example.com'],
});
```

---

### Formatting & styling

All indices are **0-based** and ranges are half-open `[start, end)`.

#### `setColumnWidth(props)`

Sets the pixel width of one or more columns.

```ts
// Set columns A–C (indices 0–2) to 150 px
await sheet.setColumnWidth({
  sheetTitle: 'Sheet1',
  startColumnIndex: 0,
  endColumnIndex: 3,
  pixelSize: 150,
});
```

---

#### `setRowHeight(props)`

Sets the pixel height of one or more rows.

```ts
// Set the header row (index 0) to 40 px
await sheet.setRowHeight({
  sheetTitle: 'Sheet1',
  startRowIndex: 0,
  endRowIndex: 1,
  pixelSize: 40,
});
```

---

#### `autoResizeColumns(props)`

Fits column widths to their content automatically.

```ts
await sheet.autoResizeColumns({
  sheetTitle: 'Report',
  startColumnIndex: 0,
  endColumnIndex: 10,
});
```

---

#### `freezeRowsAndColumns(props)`

Freezes rows and/or columns so they stay visible while scrolling. Pass `0` to unfreeze.

```ts
// Freeze the first row and first column
await sheet.freezeRowsAndColumns({
  sheetTitle: 'Data',
  frozenRowCount: 1,
  frozenColumnCount: 1,
});
```

---

#### `mergeCells(props)`

Merges a rectangular block of cells.

| `mergeType` | Behaviour |
|-------------|-----------|
| `'MERGE_ALL'` (default) | Merge entire range into one cell |
| `'MERGE_COLUMNS'` | Merge within each column independently |
| `'MERGE_ROWS'` | Merge within each row independently |

```ts
// Merge A1:C1 into a single header cell
await sheet.mergeCells({
  sheetTitle: 'Report',
  startRowIndex: 0, endRowIndex: 1,
  startColumnIndex: 0, endColumnIndex: 3,
});
```

---

#### `setCellFormat(props)`

Applies a `CellFormat` to every cell in a range. The `fields` mask controls which sub-properties are written — use a specific mask to avoid overwriting unrelated formatting.

```ts
// Bold the header row
await sheet.setCellFormat({
  sheetTitle: 'Report',
  range: 'A1:Z1',
  format: { textFormat: { bold: true } },
  fields: 'textFormat.bold',
});

// Set a red background on a range
await sheet.setCellFormat({
  sheetTitle: 'Alerts',
  range: 'B2:B50',
  format: {
    backgroundColor: { red: 1, green: 0, blue: 0 },
  },
  fields: 'backgroundColor',
});
```

---

### Data utilities

#### `getHeaderRow(sheetTitle)`

Returns the values from row 1 — typically the column headers.

```ts
const res = await sheet.getHeaderRow('Customers');
if (res.success) console.log(res.data); // ['Name', 'Email', 'Plan']
```

---

#### `updateCell(props)`

Writes a single value into one cell identified by column letter and 1-based row number.

```ts
await sheet.updateCell({
  sheetTitle: 'Tasks',
  row: 5,
  column: 'C',
  value: 'DONE',
});
```

---

#### `getDataAsObjects<T>(props)`

Reads a range and maps each data row to a typed object keyed by the header row values. This is the most ergonomic way to consume tabular data.

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `sheetTitle` | `string` | — | Target tab |
| `range` | `string` | `'A:Z'` | A1 range to read |
| `headerRow` | `number` | `1` | 1-based row number containing column headers |

```ts
type Order = { id: string; customer: string; amount: string };

const res = await sheet.getDataAsObjects<Order>({ sheetTitle: 'Orders' });
if (res.success) {
  res.data.forEach(o => console.log(o.customer, o.amount));
}
```

---

#### `copyPasteRange(props)`

Copies a range to a destination — optionally on a different sheet — using the Sheets API `copyPaste` request.

| `pasteType` | What is pasted |
|-------------|----------------|
| `'PASTE_NORMAL'` (default) | Values, formulas, and formatting |
| `'PASTE_VALUES'` | Values only (no formulas or formatting) |
| `'PASTE_FORMAT'` | Formatting only |
| `'PASTE_FORMULA'` | Formulas only |

```ts
// Copy A1:D10 values-only from 'Live' to 'Archive'
await sheet.copyPasteRange({
  sheetTitle: 'Live',
  sourceRange: 'A1:D10',
  destinationRange: 'A1:D10',
  destinationSheetTitle: 'Archive',
  pasteType: 'PASTE_VALUES',
});
```

---

### Filtering & sorting

#### `sortRange(props)`

Sorts rows within a range in-place. Multiple sort specs act as tie-breakers.
`columnIndex` is 0-based relative to the start of the range.

```ts
// Sort by column B ascending, then column C descending
await sheet.sortRange({
  sheetTitle: 'Sales',
  range: 'A2:E500', // exclude header row
  sortSpecs: [
    { columnIndex: 1 },
    { columnIndex: 2, ascending: false },
  ],
});
```

---

#### `setBasicFilter(props)`

Attaches dropdown filter controls to a range. Only one basic filter can exist per sheet; clear it first if one already exists.

```ts
await sheet.setBasicFilter({ sheetTitle: 'Inventory', range: 'A1:G500' });
```

---

#### `clearBasicFilter(sheetTitle)`

Removes the basic filter and restores all hidden rows to visibility.

```ts
await sheet.clearBasicFilter('Inventory');
```

---

### Named ranges

Named ranges provide stable, human-readable references that survive structural edits (inserted rows/columns automatically shift the range).

#### `listNamedRanges()`

Returns all named ranges defined in the spreadsheet.

```ts
const res = await sheet.listNamedRanges();
if (res.success) {
  res.data.forEach(nr => console.log(nr.name, nr.namedRangeId));
}
```

---

#### `addNamedRange(props)`

Creates a named range. The name must be unique and contain no spaces.

```ts
await sheet.addNamedRange({
  name: 'TAX_RATES',
  sheetTitle: 'Config',
  range: 'B2:B10',
});
```

---

#### `deleteNamedRange(namedRangeId)`

Deletes a named range by its API-assigned ID. The underlying cells are not affected. Use `listNamedRanges()` to find the ID.

```ts
const list = await sheet.listNamedRanges();
const id = list.data?.find(nr => nr.name === 'OLD_RANGE')?.namedRangeId;
if (id) await sheet.deleteNamedRange(id);
```

---

## TypeScript types

All prop and response types are exported from the package root:

```ts
import type {
  // Core
  GoogleSheetResponse,
  GoogleSheetClient,
  GoogleSheetTab,
  GoogleSheet,
  GoogleSheetMetadata,
  // Read
  GetDataRangeProps,
  // Write
  AppendDataProps,
  UpdateRangeProps,
  // Batch
  BatchGetRangesProps,
  BatchUpdateRangesProps,
  BatchUpdateEntry,
  BatchClearRangesProps,
  // Sheet management
  AddSheetProps,
  ProtectSheetProps,
  ProtectedRange,
  // Formatting
  SetColumnWidthProps,
  SetRowHeightProps,
  AutoResizeColumnsProps,
  FreezeProps,
  MergeCellsProps,
  CellFormatProps,
  // Data utilities
  GetDataAsObjectsProps,
  UpdateCellProps,
  CopyPasteRangeProps,
  // Filter / sort
  SortRangeProps,
  SortSpec,
  SetBasicFilterProps,
  // Named ranges
  AddNamedRangeProps,
  NamedRange,
} from '@mshindi-labs/gsheets';
```

---

## Requirements

| Requirement | Version |
|-------------|---------|
| Node.js | >= 18.0.0 |
| TypeScript | >= 5.0 (peer, optional) |
| `googleapis` | included as a dependency |

---

## License

MIT — see [LICENSE](./LICENSE).
