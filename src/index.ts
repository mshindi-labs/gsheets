// Auth
export type { GoogleAuthConfig } from './auth/types.js';
export {
  createOAuth2Client,
  generateAuthUrl,
  retrieveRefreshToken,
} from './auth/create-oauth2-client.js';

// Sheets
export type {
  GoogleSheetResponse,
  GoogleSheetClient,
  GoogleSheetTab,
  GoogleSheet,
  GoogleSheetProperties,
  GoogleSheetMetadata,
  GetDataRangeProps,
  FindRowsProps,
  AppendDataProps,
  UpdateRangeProps,
  ReturnColumnRange,
  GetRowByColumnOptions,
  Log,
  GoogleSheetRow,
  // Batch operations
  BatchGetRangesProps,
  BatchUpdateEntry,
  BatchUpdateRangesProps,
  BatchClearRangesProps,
  // Sheet management
  AddSheetProps,
  ProtectSheetProps,
  ProtectedRange,
  // Formatting & styling
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
  // Filtering & sorting
  SortSpec,
  SortRangeProps,
  SetBasicFilterProps,
  // Named ranges
  AddNamedRangeProps,
  NamedRange,
} from './sheets/types.js';
export { GoogleSpreadSheet } from './sheets/google-spreadsheet.js';
export type { GoogleSheetsConfig } from './sheets/create-sheets-client.js';
export { createSheetsClient } from './sheets/create-sheets-client.js';

// Common
export type {
  SuccessResponse,
  ErrorResponse,
  ActionResponse,
} from './common/types.js';
