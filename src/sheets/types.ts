import type { Auth, sheets_v4 } from 'googleapis';
import type { ErrorResponse, SuccessResponse } from '../common/types.js';

export type GoogleSheetResponse<T> = SuccessResponse<T> | ErrorResponse;

export type GoogleSheetClient = {
  auth: Auth.OAuth2Client;
  spreadsheetId: string;
};

export type GoogleSheetTab = {
  title?: string | null;
  sheetId?: number | null;
  index?: number | null;
  gridProperties?: sheets_v4.Schema$GridProperties | null;
};

export type GoogleSheet = sheets_v4.Schema$Sheet;
export type GoogleSheetProperties = sheets_v4.Schema$SheetProperties;
export type GoogleSheetMetadata = sheets_v4.Schema$Spreadsheet;

export type GetDataRangeProps = {
  sheetTitle?: string;
  sheetId?: number;
  range: string;
  valueRenderOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Get['valueRenderOption'];
};

export type FindRowsProps = {
  sheetTitle: string;
  column: string;
  predicate: (value: string) => boolean;
};

export type AppendDataProps = {
  sheetTitle: string;
  data: string[][];
  valueInputOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Append['valueInputOption'];
  insertDataOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Append['insertDataOption'];
};

export type UpdateRangeProps = {
  sheetTitle: string;
  range: string; // Format: 'A1', 'B3', 'C1:C10', etc.
  data: string[][];
  valueInputOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Update['valueInputOption'];
};

export type ReturnColumnRange = {
  startColumnLetter: string;
  endColumnLetter: string;
};

export type GetRowByColumnOptions = {
  sheetTitle: string;
  column: string;
  predicate: (v: unknown) => unknown;
  returnColumnRange: ReturnColumnRange;
  valueRenderOption?: string;
  requiredDataIndices?: number[];
};

export type Log = {
  message: string;
  time: string;
  index: number;
};

export type GoogleSheetRow = {
  rowIndex: number;
  entryData: string[];
};

// ---------------------------------------------------------------------------
// Batch operations
// ---------------------------------------------------------------------------

export type BatchGetRangesProps = {
  sheetTitle: string;
  ranges: string[];
  valueRenderOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Get['valueRenderOption'];
};

export type BatchUpdateEntry = {
  range: string;
  data: string[][];
};

export type BatchUpdateRangesProps = {
  sheetTitle: string;
  entries: BatchUpdateEntry[];
  valueInputOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Update['valueInputOption'];
};

export type BatchClearRangesProps = {
  sheetTitle: string;
  ranges: string[];
};

// ---------------------------------------------------------------------------
// Sheet management
// ---------------------------------------------------------------------------

export type AddSheetProps = {
  title: string;
  index?: number;
  rowCount?: number;
  columnCount?: number;
};

export type ProtectSheetProps = {
  sheetTitle: string;
  description?: string;
  warningOnly?: boolean;
  editorEmails?: string[];
};

export type ProtectedRange = sheets_v4.Schema$ProtectedRange;

// ---------------------------------------------------------------------------
// Formatting & styling
// ---------------------------------------------------------------------------

export type SetColumnWidthProps = {
  sheetTitle: string;
  startColumnIndex: number;
  endColumnIndex: number;
  pixelSize: number;
};

export type SetRowHeightProps = {
  sheetTitle: string;
  startRowIndex: number;
  endRowIndex: number;
  pixelSize: number;
};

export type AutoResizeColumnsProps = {
  sheetTitle: string;
  startColumnIndex: number;
  endColumnIndex: number;
};

export type FreezeProps = {
  sheetTitle: string;
  frozenRowCount?: number;
  frozenColumnCount?: number;
};

export type MergeCellsProps = {
  sheetTitle: string;
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
  mergeType?: 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS';
};

export type CellFormatProps = {
  sheetTitle: string;
  range: string;
  format: sheets_v4.Schema$CellFormat;
  fields?: string;
};

// ---------------------------------------------------------------------------
// Data utilities
// ---------------------------------------------------------------------------

export type GetDataAsObjectsProps = {
  sheetTitle: string;
  range?: string;
  headerRow?: number;
};

export type UpdateCellProps = {
  sheetTitle: string;
  row: number;
  column: string;
  value: string;
  valueInputOption?: sheets_v4.Params$Resource$Spreadsheets$Values$Update['valueInputOption'];
};

export type CopyPasteRangeProps = {
  sheetTitle: string;
  sourceRange: string;
  destinationRange: string;
  destinationSheetTitle?: string;
  pasteType?: 'PASTE_NORMAL' | 'PASTE_VALUES' | 'PASTE_FORMAT' | 'PASTE_FORMULA';
};

// ---------------------------------------------------------------------------
// Filtering & sorting
// ---------------------------------------------------------------------------

export type SortSpec = {
  columnIndex: number;
  ascending?: boolean;
};

export type SortRangeProps = {
  sheetTitle: string;
  range: string;
  sortSpecs: SortSpec[];
};

export type SetBasicFilterProps = {
  sheetTitle: string;
  range: string;
};

// ---------------------------------------------------------------------------
// Named ranges
// ---------------------------------------------------------------------------

export type AddNamedRangeProps = {
  name: string;
  sheetTitle: string;
  range: string;
};

export type NamedRange = sheets_v4.Schema$NamedRange;
