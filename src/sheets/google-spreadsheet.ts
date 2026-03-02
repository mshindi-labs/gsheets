import { google, sheets_v4 } from 'googleapis';
import type { ErrorResponse } from '../common/types.js';
import type {
  AddNamedRangeProps,
  AddSheetProps,
  AppendDataProps,
  AutoResizeColumnsProps,
  BatchClearRangesProps,
  BatchGetRangesProps,
  BatchUpdateRangesProps,
  CellFormatProps,
  CopyPasteRangeProps,
  FindRowsProps,
  FreezeProps,
  GetDataAsObjectsProps,
  GetDataRangeProps,
  GetRowByColumnOptions,
  GoogleSheet,
  GoogleSheetClient,
  GoogleSheetMetadata,
  GoogleSheetResponse,
  GoogleSheetRow,
  GoogleSheetTab,
  MergeCellsProps,
  NamedRange,
  ProtectedRange,
  ProtectSheetProps,
  SetBasicFilterProps,
  SetColumnWidthProps,
  SetRowHeightProps,
  SortRangeProps,
  UpdateCellProps,
  UpdateRangeProps,
} from './types.js';

// ---------------------------------------------------------------------------
// Internal error helpers
// ---------------------------------------------------------------------------

/**
 * Safely extracts a human-readable message from any thrown value.
 * The Sheets API throws GaxiosError objects whose `.message` property
 * contains the HTTP status and description — this unwraps that string
 * rather than producing "[object Object]".
 */
function getErrorMessage(error: unknown): string {
  if (typeof error === 'object' && error !== null && 'message' in error) {
    return String((error as { message: unknown }).message);
  }
  return String(error);
}

/**
 * Wraps a caught error into a typed {@link ErrorResponse} so every
 * `catch` block can return a consistent shape instead of re-throwing.
 *
 * @param context - Human-readable description of the operation that failed.
 * @param error   - The raw caught value.
 */
function handleError(context: string, error: unknown): ErrorResponse {
  return {
    success: false,
    problem: `${context} Error: ${getErrorMessage(error)}`,
  };
}

// ---------------------------------------------------------------------------
// GoogleSpreadSheet class
// ---------------------------------------------------------------------------

/**
 * A high-level, fully-typed wrapper around the Google Sheets v4 API.
 *
 * Every method returns a discriminated-union {@link GoogleSheetResponse}:
 * - `{ success: true,  data: T }` on success
 * - `{ success: false, problem: string }` on failure
 *
 * This means callers never need to catch exceptions — simply check
 * `result.success` before accessing `result.data`.
 *
 * @example
 * ```ts
 * import { createSheetsClient } from '@mshindi-labs/gsheets';
 *
 * const sheet = createSheetsClient(
 *   { clientId, clientSecret, redirectUri, refreshToken },
 *   'YOUR_SPREADSHEET_ID',
 * );
 *
 * const result = await sheet.getHeaderRow('Sheet1');
 * if (result.success) console.log(result.data); // string[]
 * ```
 */
export class GoogleSpreadSheet {
  private readonly spreadsheetId: string;
  private readonly google_sheets: sheets_v4.Sheets;

  constructor({ auth, spreadsheetId }: GoogleSheetClient) {
    this.spreadsheetId = spreadsheetId;
    // Instantiate the low-level googleapis client once and reuse it across
    // all method calls to avoid recreating the HTTP transport repeatedly.
    this.google_sheets = google.sheets({ version: 'v4', auth });
  }

  /**
   * Returns `true` for any value that is non-null, non-undefined, and not
   * an empty/whitespace-only string. Used when checking whether a cell or
   * column value actually contains meaningful data.
   */
  private isTruthy(value: unknown): boolean {
    if (value === null || value === undefined) return false;
    if (typeof value === 'string' && value.trim() === '') return false;
    return true;
  }

  /**
   * Exposes the raw `sheets_v4.Sheets` client for one-off API calls that
   * are not yet covered by this wrapper.
   *
   * @returns The underlying googleapis Sheets client.
   */
  async _document(): Promise<sheets_v4.Sheets> {
    return this.google_sheets;
  }

  // -------------------------------------------------------------------------
  // Metadata / structure
  // -------------------------------------------------------------------------

  /**
   * Fetches the spreadsheet's top-level properties together with the
   * properties of every sheet it contains.
   *
   * Useful for reading the spreadsheet title, locale, time zone, and
   * discovering how many sheets exist without fetching cell data.
   *
   * @returns The full `Spreadsheet` resource from the Sheets API.
   *
   * @example
   * ```ts
   * const meta = await sheet.getMetadata();
   * if (meta.success) console.log(meta.data.properties?.title);
   * ```
   */
  async getMetadata(): Promise<GoogleSheetResponse<GoogleSheetMetadata>> {
    try {
      const response = await this.google_sheets.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
        // Request only the fields we care about to reduce response payload.
        fields: 'properties,sheets.properties',
      });

      if (!response.data) {
        return {
          success: false,
          problem: `No metadata found for spreadsheet [${this.spreadsheetId}]`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Get metadata of the spreadsheet [${this.spreadsheetId}]`,
        error,
      );
    }
  }

  /**
   * Returns a lightweight summary of every sheet (tab) in the spreadsheet:
   * title, numeric sheetId, tab index, and grid dimensions.
   *
   * This is the cheapest way to discover available sheets and resolve a
   * human-readable title to the numeric `sheetId` required by batchUpdate
   * requests.
   *
   * @returns Array of {@link GoogleSheetTab} objects, one per tab.
   *
   * @example
   * ```ts
   * const tabs = await sheet.getTabs();
   * if (tabs.success) {
   *   tabs.data.forEach(t => console.log(t.title, t.sheetId));
   * }
   * ```
   */
  async getTabs(): Promise<GoogleSheetResponse<GoogleSheetTab[]>> {
    try {
      const { data: metadata } = await this.google_sheets.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
        fields: 'sheets.properties',
      });

      const tabs: GoogleSheetTab[] =
        metadata.sheets?.map((sheet) => {
          const { title, sheetId, index, gridProperties } =
            sheet.properties ?? {};
          return { title, sheetId, index, gridProperties };
        }) ?? [];

      return { success: true, data: tabs };
    } catch (error) {
      return handleError(
        `Get tabs of the spreadsheet [${this.spreadsheetId}]`,
        error,
      );
    }
  }

  /**
   * Finds a sheet by its human-readable title and returns the full
   * `Sheet` resource, including its properties and (if requested) data.
   *
   * @param sheetTitle - The exact title as it appears on the tab.
   * @returns The matching `Sheet` resource, or a failure if not found.
   *
   * @example
   * ```ts
   * const result = await sheet.getSheetByName('Inventory');
   * if (result.success) {
   *   console.log(result.data.properties?.sheetId);
   * }
   * ```
   */
  async getSheetByName(
    sheetTitle: string,
  ): Promise<GoogleSheetResponse<GoogleSheet>> {
    try {
      const metadata = await this.getMetadata();
      if (!metadata.success) {
        return {
          success: false,
          problem: 'Unable to retrieve metadata for spreadsheet.',
        };
      }

      const sheet = metadata.data.sheets?.find(
        (s) => s.properties?.title === sheetTitle,
      );
      if (!sheet) {
        return { success: false, problem: `Sheet ${sheetTitle} not found.` };
      }
      return { success: true, data: sheet };
    } catch (error) {
      return handleError(`Error fetching sheet by name ${sheetTitle}.`, error);
    }
  }

  // -------------------------------------------------------------------------
  // Read operations
  // -------------------------------------------------------------------------

  /**
   * Reads a rectangular range of cells and returns the raw API
   * `ValueRange` object (includes the range address and values matrix).
   *
   * Accepts either a human-readable `sheetTitle` or a numeric `sheetId`;
   * if only `sheetId` is given the method resolves the title first.
   *
   * @param props - {@link GetDataRangeProps}
   * @param props.sheetTitle - Title of the target sheet (preferred).
   * @param props.sheetId    - Numeric sheet ID (used when title is unknown).
   * @param props.range      - A1 notation range, e.g. `'A1:D20'`.
   * @param props.valueRenderOption - How values are rendered; defaults to `'FORMATTED_VALUE'`.
   * @returns The `ValueRange` resource containing the cell data.
   *
   * @example
   * ```ts
   * const res = await sheet.getDataRange({ sheetTitle: 'Orders', range: 'A1:E100' });
   * if (res.success) console.log(res.data.values);
   * ```
   */
  async getDataRange(
    props: GetDataRangeProps,
  ): Promise<GoogleSheetResponse<sheets_v4.Schema$ValueRange>> {
    const { sheetTitle, sheetId, range, valueRenderOption } = props;

    if (!sheetTitle && !sheetId) {
      return {
        success: false,
        problem: 'Either sheetTitle or sheetId must be provided.',
      };
    }

    let finalSheetTitle: string | undefined = sheetTitle;

    // When only a numeric sheetId is available, look up the matching title
    // so we can build the "SheetTitle!Range" address string the API expects.
    if (!sheetTitle && sheetId) {
      const tabsResponse = await this.getTabs();
      if (!tabsResponse.success) {
        return {
          success: false,
          problem: `Failed to fetch tabs: ${tabsResponse.problem}`,
        };
      }

      const matchingSheet = tabsResponse.data.find(
        (tab) => tab.sheetId === sheetId,
      );
      if (!matchingSheet) {
        return {
          success: false,
          problem: `No sheet found with the provided sheetId: ${sheetId}`,
        };
      }

      finalSheetTitle = matchingSheet.title ?? undefined;
    }

    try {
      const response = await this.google_sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${finalSheetTitle}!${range}`,
        valueRenderOption: valueRenderOption ?? 'FORMATTED_VALUE',
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to fetch data range for sheet ${finalSheetTitle}`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Error fetching data range for sheet ${finalSheetTitle}.`,
        error,
      );
    }
  }

  /**
   * Returns the 1-based row index of the last row that contains data in
   * a given column (defaults to column A when none is specified).
   *
   * Internally reads the entire column and counts non-empty entries, which
   * is an efficient way to find the end of a dataset without knowing its
   * size in advance.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param column     - Column letter, e.g. `'B'`. Defaults to `'A'`.
   * @returns The row count (= last row with data), or `0` if the column is empty.
   *
   * @example
   * ```ts
   * const res = await sheet.getLastRowWithData('Inventory', 'C');
   * if (res.success) console.log('Last row:', res.data); // e.g. 42
   * ```
   */
  async getLastRowWithData(
    sheetTitle: string,
    column?: string,
  ): Promise<GoogleSheetResponse<number>> {
    const range = column ? `${column}:${column}` : 'A:A';

    try {
      const response = await this.google_sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetTitle}!${range}`,
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to fetch data for sheet ${sheetTitle}`,
        };
      }

      return { success: true, data: response.data.values?.length ?? 0 };
    } catch (error) {
      return handleError(`Error fetching data for sheet ${sheetTitle}.`, error);
    }
  }

  /**
   * Reads all cell values from a single row and returns them as a flat
   * string array. Row numbers are 1-based (row 1 is the first row).
   *
   * @param params.sheetTitle         - Title of the target sheet.
   * @param params.rowNumber          - 1-based row number to read.
   * @param params.valueRenderOption  - Defaults to `'UNFORMATTED_VALUE'`.
   * @returns Array of cell values in column order, empty array if row is blank.
   *
   * @example
   * ```ts
   * const row = await sheet.getRow({ sheetTitle: 'Sales', rowNumber: 5 });
   * if (row.success) console.log(row.data); // ['Alice', '2024-01-15', '1200']
   * ```
   */
  async getRow({
    sheetTitle,
    rowNumber,
    valueRenderOption = 'UNFORMATTED_VALUE',
  }: {
    sheetTitle: string;
    rowNumber: number;
    valueRenderOption?: string;
  }): Promise<GoogleSheetResponse<string[]>> {
    try {
      const response = await this.google_sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        // "5:5" reads the entire 5th row across all columns.
        range: `${sheetTitle}!${rowNumber}:${rowNumber}`,
        valueRenderOption,
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to fetch data for row ${rowNumber} in sheet ${sheetTitle}`,
        };
      }

      // The API returns a nested array; the first (and only) sub-array is our row.
      return { success: true, data: response.data.values?.[0] ?? [] };
    } catch (error) {
      return handleError(
        `Error fetching data for row ${rowNumber} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  /**
   * Reads a range from a named sheet and returns the raw 2-D values array.
   * Equivalent to `getDataRange` but returns only `values` rather than the
   * full `ValueRange` resource.
   *
   * @param sheetTitle          - Title of the target sheet.
   * @param range               - A1 notation, e.g. `'B2:F50'`.
   * @param valueRenderOption   - Defaults to `'UNFORMATTED_VALUE'`.
   * @returns 2-D array of cell values, or `null` if the range is empty.
   *
   * @example
   * ```ts
   * const res = await sheet.getValuesBySheetTitle('Config', 'A1:B20');
   * if (res.success) console.log(res.data);
   * ```
   */
  async getValuesBySheetTitle(
    sheetTitle: string,
    range: string,
    valueRenderOption = 'UNFORMATTED_VALUE',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ): Promise<GoogleSheetResponse<any>> {
    try {
      const response = await this.google_sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetTitle}!${range}`,
        valueRenderOption,
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to fetch data for range ${range} in sheet ${sheetTitle}`,
        };
      }

      return { success: true, data: response.data.values };
    } catch (error) {
      return handleError(
        `Error fetching data for range ${range} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  /**
   * Reads every cell in a single column and returns the values as a 2-D
   * array (each element is a one-item array representing a row).
   *
   * This is a thin wrapper around {@link getValuesBySheetTitle} that
   * constructs the full-column range string (e.g. `'C:C'`) automatically.
   *
   * @param sheetTitle         - Title of the target sheet.
   * @param column             - Column letter(s), e.g. `'A'` or `'AA'`.
   * @param valueRenderOption  - Defaults to `'UNFORMATTED_VALUE'`.
   * @returns 2-D array where each row is `[cellValue]`.
   *
   * @example
   * ```ts
   * const col = await sheet.getColumnValues('Employees', 'B');
   * if (col.success) col.data.forEach(([name]) => console.log(name));
   * ```
   */
  async getColumnValues(
    sheetTitle: string,
    column: string,
    valueRenderOption = 'UNFORMATTED_VALUE',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ): Promise<GoogleSheetResponse<any>> {
    return this.getValuesBySheetTitle(
      sheetTitle,
      `${column}:${column}`,
      valueRenderOption,
    );
  }

  // -------------------------------------------------------------------------
  // Search / query
  // -------------------------------------------------------------------------

  /**
   * Scans a column and returns the 1-based row indices of every cell that
   * satisfies the provided predicate.
   *
   * Useful when you need to locate rows by a known value (e.g. find all rows
   * where status is `'PENDING'`) before performing follow-up reads or writes.
   *
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.column     - Column letter to scan, e.g. `'C'`.
   * @param props.predicate  - Function that receives each cell value and returns `true` to include that row.
   * @returns Array of 1-based row numbers whose cells matched the predicate.
   *
   * @example
   * ```ts
   * const res = await sheet.findRowsMatchingCriteria({
   *   sheetTitle: 'Orders',
   *   column: 'D',
   *   predicate: (v) => v === 'SHIPPED',
   * });
   * if (res.success) console.log('Shipped rows:', res.data);
   * ```
   */
  async findRowsMatchingCriteria({
    sheetTitle,
    column,
    predicate,
  }: FindRowsProps): Promise<GoogleSheetResponse<number[]>> {
    try {
      const response = await this.getDataRange({
        sheetTitle,
        range: `${column}:${column}`,
      });

      if (!response.success || !response.data.values) {
        return {
          success: false,
          problem: `Failed to fetch data for column ${column} in sheet ${sheetTitle}`,
        };
      }

      const matchingRows: number[] = [];
      response.data.values.forEach((row, index) => {
        const cellValue = row[0] as string;
        if (predicate(cellValue)) {
          // Convert 0-based index to 1-based row number.
          matchingRows.push(index + 1);
        }
      });

      return { success: true, data: matchingRows };
    } catch (error) {
      return handleError(
        `Error searching for rows in column ${column} of sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  /**
   * Finds rows where a specific column satisfies a predicate, then fetches
   * the full row data across a configurable column range for each match.
   *
   * This two-phase approach (find → fetch) is appropriate when you know
   * which column to key on but need to read adjacent columns once a match
   * is confirmed. An optional `requiredDataIndices` filter lets you discard
   * rows that are structurally incomplete (missing required columns).
   *
   * @param options - {@link GetRowByColumnOptions}
   * @param options.sheetTitle         - Title of the target sheet.
   * @param options.column             - Key column to scan.
   * @param options.predicate          - Matching function for each key-column cell.
   * @param options.returnColumnRange  - Start/end column letters of data to return per matched row.
   * @param options.valueRenderOption  - Optional render mode; defaults to `'UNFORMATTED_VALUE'`.
   * @param options.requiredDataIndices - 0-based column indices within `returnColumnRange` that must be non-empty.
   * @returns Array of `{ rowIndex, entryData }` for every qualifying row.
   *
   * @example
   * ```ts
   * const res = await sheet.findEntryRowByColumValue({
   *   sheetTitle: 'Invoices',
   *   column: 'A',
   *   predicate: (v) => v === 'INV-007',
   *   returnColumnRange: { startColumnLetter: 'A', endColumnLetter: 'F' },
   * });
   * if (res.success) console.log(res.data[0].entryData);
   * ```
   */
  async findEntryRowByColumValue(
    options: GetRowByColumnOptions,
  ): Promise<GoogleSheetResponse<GoogleSheetRow[]>> {
    const {
      sheetTitle,
      column,
      predicate,
      returnColumnRange,
      valueRenderOption,
      requiredDataIndices,
    } = options;

    const _valueRenderOption = valueRenderOption ?? 'UNFORMATTED_VALUE';

    const response = await this.getColumnValues(
      sheetTitle,
      column,
      _valueRenderOption,
    );

    if (!response.success) {
      return response;
    }

    const columnData = response.data as Array<Array<string>>;

    try {
      // Map every cell to an async task that either resolves to the full row
      // data or `false` (skipped). Running all tasks concurrently with
      // Promise.all is more efficient than sequential awaits when the sheet
      // has many rows to check.
      const getRowIndices = columnData.map(
        async (value: Array<string>, index: number) => {
          const hasValue = this.isTruthy(value[0]);
          if (!hasValue) return false;
          if (!predicate(value[0])) return false;

          // Row numbers in the API are 1-based.
          const rowIndex = index + 1;

          const entryRowResponse = await this.getValuesBySheetTitle(
            sheetTitle,
            `${returnColumnRange.startColumnLetter}${rowIndex}:${returnColumnRange.endColumnLetter}${rowIndex}`,
            _valueRenderOption,
          );

          if (!entryRowResponse.success) return false;

          const entryRowData = (entryRowResponse.data as unknown[][] | null | undefined) ?? [[]];
          const rowData = entryRowData[0] ?? [];

          if (!requiredDataIndices) {
            return { rowIndex, entryData: rowData };
          }

          // Only return the row if every required column index has a truthy value.
          const meetsAllCriteria = requiredDataIndices.every((i) =>
            this.isTruthy(rowData[i]),
          );

          if (!meetsAllCriteria) return false;

          return { rowIndex, entryData: rowData };
        },
      );

      const rowIndices = await Promise.all(getRowIndices);
      const data = rowIndices.filter(
        (r): r is GoogleSheetRow => r !== false,
      );

      return { success: true, data };
    } catch (error) {
      return handleError(
        `Error fetching data for column ${column} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Write operations
  // -------------------------------------------------------------------------

  /**
   * Appends one or more rows of data to the end of the sheet's existing
   * dataset. The API automatically finds the first empty row after the
   * last row that contains data.
   *
   * `insertDataOption: 'INSERT_ROWS'` (the default) inserts new rows rather
   * than overwriting the first empty row, which is the safer default for
   * most append use cases.
   *
   * @param props - {@link AppendDataProps}
   * @param props.sheetTitle        - Title of the target sheet.
   * @param props.data              - 2-D array of values to append.
   * @param props.valueInputOption  - `'USER_ENTERED'` (default) or `'RAW'`.
   * @param props.insertDataOption  - `'INSERT_ROWS'` (default) or `'OVERWRITE'`.
   * @returns The `AppendValuesResponse` with update metadata.
   *
   * @example
   * ```ts
   * await sheet.appendData({
   *   sheetTitle: 'Log',
   *   data: [['2024-06-01', 'Deploy', 'success']],
   * });
   * ```
   */
  async appendData({
    sheetTitle,
    data,
    valueInputOption = 'USER_ENTERED',
    insertDataOption = 'INSERT_ROWS',
  }: AppendDataProps): Promise<
    GoogleSheetResponse<sheets_v4.Schema$AppendValuesResponse>
  > {
    try {
      const response = await this.google_sheets.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: sheetTitle,
        valueInputOption,
        insertDataOption,
        requestBody: { values: data },
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to append data to sheet ${sheetTitle}`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(`Error appending data to sheet ${sheetTitle}.`, error);
    }
  }

  /**
   * Writes data into a specific range, overwriting any existing values.
   * Use this when you know the exact target cells (e.g. refreshing a
   * summary block) rather than appending to the end of a table.
   *
   * @param props - {@link UpdateRangeProps}
   * @param props.sheetTitle       - Title of the target sheet.
   * @param props.range            - A1 notation start cell or range, e.g. `'B3'` or `'C1:C10'`.
   * @param props.data             - 2-D array of new values.
   * @param props.valueInputOption - `'USER_ENTERED'` (default) or `'RAW'`.
   * @returns The `UpdateValuesResponse` containing the number of cells updated.
   *
   * @example
   * ```ts
   * await sheet.updateRange({
   *   sheetTitle: 'Summary',
   *   range: 'B2',
   *   data: [['Updated value']],
   * });
   * ```
   */
  async updateRange({
    sheetTitle,
    range,
    data,
    valueInputOption = 'USER_ENTERED',
  }: UpdateRangeProps): Promise<
    GoogleSheetResponse<sheets_v4.Schema$UpdateValuesResponse>
  > {
    try {
      const response = await this.google_sheets.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetTitle}!${range}`,
        valueInputOption,
        requestBody: { values: data },
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to update range ${range} in sheet ${sheetTitle}`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Error updating range ${range} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  /**
   * Erases all values in a given range while leaving formatting intact.
   * Equivalent to selecting cells in the UI and pressing Delete.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param range      - A1 notation range to clear, e.g. `'A2:Z100'`.
   *
   * @example
   * ```ts
   * await sheet.clearRange('Temp', 'A1:Z1000');
   * ```
   */
  async clearRange(
    sheetTitle: string,
    range: string,
  ): Promise<GoogleSheetResponse<void>> {
    try {
      await this.google_sheets.spreadsheets.values.clear({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetTitle}!${range}`,
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error clearing range ${range} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Dimension operations (delete / insert rows & columns)
  // -------------------------------------------------------------------------

  /**
   * Deletes a contiguous block of rows from a sheet.
   * Indices are 0-based and the range is half-open `[startIndex, endIndex)`,
   * matching the Sheets API convention.
   *
   * **Warning:** this operation is irreversible — deleted rows cannot be
   * recovered programmatically.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param startIndex - 0-based index of the first row to delete.
   * @param endIndex   - 0-based index one past the last row to delete.
   *
   * @example
   * ```ts
   * // Delete rows 2–4 (0-based indices 1–3, endIndex exclusive = 4)
   * await sheet.deleteRows('Sheet1', 1, 4);
   * ```
   */
  async deleteRows(
    sheetTitle: string,
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    const tabsResponse = await this.getTabs();
    if (!tabsResponse.success) {
      return {
        success: false,
        problem: `Failed to fetch tabs: ${tabsResponse.problem}`,
      };
    }

    const matchingSheet = tabsResponse.data.find(
      (tab) => tab.title === sheetTitle,
    );
    const sheetId = matchingSheet?.sheetId ?? undefined;

    return this.deleteDimension(sheetId, 'ROWS', startIndex, endIndex);
  }

  /**
   * Deletes a contiguous block of columns from a sheet.
   * Indices are 0-based and the range is half-open `[startIndex, endIndex)`.
   *
   * **Warning:** this operation is irreversible.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param startIndex - 0-based index of the first column to delete.
   * @param endIndex   - 0-based index one past the last column to delete.
   *
   * @example
   * ```ts
   * // Delete columns C and D (0-based: 2 to 4)
   * await sheet.deleteColumns('Sheet1', 2, 4);
   * ```
   */
  async deleteColumns(
    sheetTitle: string,
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    const tabsResponse = await this.getTabs();
    if (!tabsResponse.success) {
      return {
        success: false,
        problem: `Failed to fetch tabs: ${tabsResponse.problem}`,
      };
    }

    const matchingSheet = tabsResponse.data.find(
      (tab) => tab.title === sheetTitle,
    );
    const sheetId = matchingSheet?.sheetId ?? undefined;

    return this.deleteDimension(sheetId, 'COLUMNS', startIndex, endIndex);
  }

  /**
   * Shared implementation for {@link deleteRows} and {@link deleteColumns}.
   * Sends a single `deleteDimension` batchUpdate request.
   */
  private async deleteDimension(
    sheetId: number | undefined,
    dimension: 'ROWS' | 'COLUMNS',
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: { sheetId, dimension, startIndex, endIndex },
              },
            },
          ],
        },
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to delete ${dimension} in sheet ${sheetId}`,
        };
      }

      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error deleting ${dimension} in sheet ${sheetId}.`,
        error,
      );
    }
  }

  /**
   * Inserts blank rows at the specified position, shifting existing rows down.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param startIndex - 0-based index where the first new row will be inserted.
   * @param endIndex   - 0-based exclusive end; determines how many rows are added (`endIndex - startIndex`).
   *
   * @example
   * ```ts
   * // Insert 3 blank rows before the current row 2 (0-based index 1)
   * await sheet.insertRows('Sheet1', 1, 4);
   * ```
   */
  async insertRows(
    sheetTitle: string,
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    return this.insertDimension(sheetTitle, 'ROWS', startIndex, endIndex);
  }

  /**
   * Inserts blank columns at the specified position, shifting existing columns right.
   *
   * @param sheetTitle - Title of the target sheet.
   * @param startIndex - 0-based index where the first new column will be inserted.
   * @param endIndex   - 0-based exclusive end.
   *
   * @example
   * ```ts
   * // Insert 1 blank column before column B (0-based index 1)
   * await sheet.insertColumns('Sheet1', 1, 2);
   * ```
   */
  async insertColumns(
    sheetTitle: string,
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    return this.insertDimension(sheetTitle, 'COLUMNS', startIndex, endIndex);
  }

  /**
   * Shared implementation for {@link insertRows} and {@link insertColumns}.
   * Resolves the numeric sheetId from the title, then sends an
   * `insertDimension` batchUpdate request.
   *
   * `inheritFromBefore: false` ensures newly inserted rows/columns receive
   * default formatting rather than copying the style of adjacent cells.
   */
  private async insertDimension(
    sheetTitle: string,
    dimension: 'ROWS' | 'COLUMNS',
    startIndex: number,
    endIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    const tabsResponse = await this.getTabs();
    if (!tabsResponse.success) {
      return {
        success: false,
        problem: `Failed to fetch tabs: ${tabsResponse.problem}`,
      };
    }

    const matchingSheet = tabsResponse.data.find(
      (tab) => tab.title === sheetTitle,
    );
    const sheetId = matchingSheet?.sheetId ?? undefined;

    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: { sheetId, dimension, startIndex, endIndex },
                inheritFromBefore: false,
              },
            },
          ],
        },
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to insert ${dimension} in sheet ${sheetTitle}`,
        };
      }

      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error inserting ${dimension} in sheet ${sheetTitle}.`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Sheet management
  // -------------------------------------------------------------------------

  /**
   * Creates a full copy of an existing sheet within the same spreadsheet.
   * The new sheet receives all data, formatting, and formulas from the source.
   *
   * @param sheetId  - Numeric ID of the source sheet (use {@link getTabs} to look it up).
   * @param newTitle - Optional title for the duplicate; the API appends `'Copy of ...'` if omitted.
   * @returns The `Sheet` resource for the newly created duplicate.
   *
   * @example
   * ```ts
   * const tabs = await sheet.getTabs();
   * const templateId = tabs.data.find(t => t.title === 'Template')?.sheetId;
   * await sheet.duplicateSheet(templateId!, 'January Report');
   * ```
   */
  async duplicateSheet(
    sheetId: number,
    newTitle?: string,
  ): Promise<GoogleSheetResponse<GoogleSheet>> {
    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              duplicateSheet: {
                sourceSheetId: sheetId,
                newSheetName: newTitle,
              },
            },
          ],
        },
      });

      const duplicated = response.data?.replies?.[0]?.duplicateSheet;
      if (!duplicated) {
        return {
          success: false,
          problem: `Failed to duplicate sheet with ID ${sheetId}`,
        };
      }

      return { success: true, data: duplicated };
    } catch (error) {
      return handleError(
        `Error duplicating sheet with ID ${sheetId}.`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Private helpers
  // -------------------------------------------------------------------------

  /**
   * Resolves the numeric `sheetId` for a given sheet title by calling
   * {@link getTabs}. Many Sheets API requests (batchUpdate, formatting,
   * dimension operations) require the numeric ID rather than the title.
   *
   * Returns a typed failure instead of throwing so callers can propagate the
   * error using a simple early-return pattern.
   */
  private async _getSheetId(
    sheetTitle: string,
  ): Promise<GoogleSheetResponse<number>> {
    const tabsResponse = await this.getTabs();
    if (!tabsResponse.success) {
      return {
        success: false,
        problem: `Failed to fetch tabs: ${tabsResponse.problem}`,
      };
    }
    const match = tabsResponse.data.find((t) => t.title === sheetTitle);
    if (!match || match.sheetId == null) {
      return { success: false, problem: `Sheet "${sheetTitle}" not found.` };
    }
    return { success: true, data: match.sheetId };
  }

  /**
   * Converts a column letter string (e.g. `'A'`, `'Z'`, `'AA'`, `'BC'`)
   * to a 0-based column index.
   *
   * The algorithm treats the string as a bijective base-26 number:
   * A=1, B=2, …, Z=26, AA=27, AB=28 … then subtracts 1 to make it 0-based.
   * This matches how the Sheets API numbers columns in `GridRange`.
   */
  private _colToIndex(col: string): number {
    let v = 0;
    for (const ch of col.toUpperCase()) {
      // charCode of 'A' is 65, so `ch.charCodeAt(0) - 64` gives A=1, B=2, …
      v = v * 26 + (ch.charCodeAt(0) - 64);
    }
    // Convert bijective base-26 result to 0-based index.
    return v - 1;
  }

  /**
   * Converts an A1 notation range string into a `GridRange` object suitable
   * for use in batchUpdate requests (formatting, sorting, filtering, etc.).
   *
   * Handles all common formats:
   * - Full range:         `'B2:D10'`  → `{ startRowIndex:1, endRowIndex:10, startColumnIndex:1, endColumnIndex:4 }`
   * - Single cell:        `'C5'`      → `{ startRowIndex:4, endRowIndex:5, startColumnIndex:2, endColumnIndex:3 }`
   * - Full column:        `'A:A'`     → `{ startColumnIndex:0, endColumnIndex:1 }` (no row bounds)
   * - Full row:           `'3:3'`     → `{ startRowIndex:2, endRowIndex:3 }` (no column bounds)
   *
   * All indices are 0-based; end indices are exclusive (+1), per the API spec.
   *
   * @param sheetId - Numeric sheet ID to embed in the `GridRange`.
   * @param range   - A1 notation string (may or may not include a colon).
   */
  private _a1ToGridRange(
    sheetId: number,
    range: string,
  ): sheets_v4.Schema$GridRange {
    const [startRef, endRef] = range.split(':');

    // Parse a single cell reference like "B2", "AA", or "5" into optional
    // column and row components.
    const parseRef = (ref: string) => {
      const m = ref.match(/^([A-Za-z]*)(\d*)$/);
      if (!m) return {};
      return {
        col: m[1] ? this._colToIndex(m[1]) : undefined,
        row: m[2] ? parseInt(m[2], 10) - 1 : undefined, // API uses 0-based rows
      };
    };

    const s = parseRef(startRef);
    const e = endRef ? parseRef(endRef) : undefined;

    const gr: sheets_v4.Schema$GridRange = { sheetId };

    if (s.col !== undefined) gr.startColumnIndex = s.col;
    if (s.row !== undefined) gr.startRowIndex = s.row;

    if (e) {
      // End indices in GridRange are exclusive, so add 1 to the parsed values.
      if (e.col !== undefined) gr.endColumnIndex = e.col + 1;
      if (e.row !== undefined) gr.endRowIndex = e.row + 1;
    } else {
      // No colon — single cell; end = start + 1 to form a 1×1 grid range.
      if (s.col !== undefined) gr.endColumnIndex = s.col + 1;
      if (s.row !== undefined) gr.endRowIndex = s.row + 1;
    }

    return gr;
  }

  /**
   * Shared implementation for {@link hideSheet} and {@link showSheet}.
   * Sends an `updateSheetProperties` batchUpdate with `fields: 'hidden'` so
   * only the visibility flag is touched and no other properties are reset.
   */
  private async _setSheetVisibility(
    sheetTitle: string,
    hidden: boolean,
  ): Promise<GoogleSheetResponse<void>> {
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId, hidden },
                // The `fields` mask is mandatory: without it the API would
                // overwrite all other sheet properties with zero values.
                fields: 'hidden',
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error ${hidden ? 'hiding' : 'showing'} sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Shared implementation for {@link setColumnWidth} and {@link setRowHeight}.
   * Sends an `updateDimensionProperties` batchUpdate that sets only
   * `pixelSize`, leaving all other dimension properties (e.g. `hiddenByUser`)
   * untouched thanks to the `fields: 'pixelSize'` mask.
   */
  private async _setDimensionPixelSize(
    sheetId: number,
    dimension: 'COLUMNS' | 'ROWS',
    start: number,
    end: number,
    pixelSize: number,
  ): Promise<GoogleSheetResponse<void>> {
    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateDimensionProperties: {
                range: { sheetId, dimension, startIndex: start, endIndex: end },
                properties: { pixelSize },
                fields: 'pixelSize',
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error setting ${dimension} pixel size in sheet ${sheetId}.`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Batch operations
  // -------------------------------------------------------------------------

  /**
   * Fetches multiple ranges in a single API call, drastically reducing
   * quota consumption and round-trip latency compared to issuing one
   * `values.get` per range.
   *
   * All ranges are automatically prefixed with `${sheetTitle}!` so you
   * only need to pass the bare A1 notation (e.g. `'A1:B10'`).
   *
   * @param props - {@link BatchGetRangesProps}
   * @param props.sheetTitle          - Title of the target sheet.
   * @param props.ranges              - Array of A1 notation ranges to fetch.
   * @param props.valueRenderOption   - Defaults to `'FORMATTED_VALUE'`.
   * @returns `BatchGetValuesResponse` containing one `ValueRange` per requested range.
   *
   * @example
   * ```ts
   * const res = await sheet.batchGetRanges({
   *   sheetTitle: 'Dashboard',
   *   ranges: ['A1:A10', 'C1:C10', 'E1:E10'],
   * });
   * if (res.success) res.data.valueRanges?.forEach(vr => console.log(vr.values));
   * ```
   */
  async batchGetRanges(
    props: BatchGetRangesProps,
  ): Promise<GoogleSheetResponse<sheets_v4.Schema$BatchGetValuesResponse>> {
    const { sheetTitle, ranges, valueRenderOption } = props;
    try {
      const response = await this.google_sheets.spreadsheets.values.batchGet({
        spreadsheetId: this.spreadsheetId,
        ranges: ranges.map((r) => `${sheetTitle}!${r}`),
        valueRenderOption: valueRenderOption ?? 'FORMATTED_VALUE',
      });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to batch get ranges in sheet "${sheetTitle}"`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Error batch getting ranges in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Writes to multiple ranges in a single API call. Each entry specifies
   * its own range and data matrix, so you can update non-contiguous areas
   * of a sheet atomically.
   *
   * @param props - {@link BatchUpdateRangesProps}
   * @param props.sheetTitle       - Title of the target sheet.
   * @param props.entries          - Array of `{ range, data }` pairs.
   * @param props.valueInputOption - `'USER_ENTERED'` (default) or `'RAW'`.
   * @returns `BatchUpdateValuesResponse` with per-range update metadata.
   *
   * @example
   * ```ts
   * await sheet.batchUpdateRanges({
   *   sheetTitle: 'Config',
   *   entries: [
   *     { range: 'B2', data: [['v1.2.0']] },
   *     { range: 'B3', data: [['2024-06-01']] },
   *   ],
   * });
   * ```
   */
  async batchUpdateRanges(
    props: BatchUpdateRangesProps,
  ): Promise<
    GoogleSheetResponse<sheets_v4.Schema$BatchUpdateValuesResponse>
  > {
    const { sheetTitle, entries, valueInputOption } = props;
    try {
      const response =
        await this.google_sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: this.spreadsheetId,
          requestBody: {
            valueInputOption: valueInputOption ?? 'USER_ENTERED',
            data: entries.map(({ range, data }) => ({
              range: `${sheetTitle}!${range}`,
              values: data,
            })),
          },
        });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to batch update ranges in sheet "${sheetTitle}"`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Error batch updating ranges in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Clears values from multiple ranges in a single API call.
   * Formatting is preserved; only cell values are erased.
   *
   * @param props - {@link BatchClearRangesProps}
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.ranges     - Array of A1 notation ranges to clear.
   * @returns `BatchClearValuesResponse` listing the cleared ranges.
   *
   * @example
   * ```ts
   * await sheet.batchClearRanges({
   *   sheetTitle: 'Cache',
   *   ranges: ['A2:A100', 'D2:D100'],
   * });
   * ```
   */
  async batchClearRanges(
    props: BatchClearRangesProps,
  ): Promise<GoogleSheetResponse<sheets_v4.Schema$BatchClearValuesResponse>> {
    const { sheetTitle, ranges } = props;
    try {
      const response =
        await this.google_sheets.spreadsheets.values.batchClear({
          spreadsheetId: this.spreadsheetId,
          requestBody: {
            ranges: ranges.map((r) => `${sheetTitle}!${r}`),
          },
        });

      if (!response.data) {
        return {
          success: false,
          problem: `Failed to batch clear ranges in sheet "${sheetTitle}"`,
        };
      }

      return { success: true, data: response.data };
    } catch (error) {
      return handleError(
        `Error batch clearing ranges in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Extended sheet management
  // -------------------------------------------------------------------------

  /**
   * Creates a new, empty sheet (tab) within the spreadsheet.
   *
   * @param props - {@link AddSheetProps}
   * @param props.title       - Title for the new sheet. Must be unique within the spreadsheet.
   * @param props.index       - 0-based position in the tab bar. Appended at the end if omitted.
   * @param props.rowCount    - Initial row count (default: 1000).
   * @param props.columnCount - Initial column count (default: 26).
   * @returns The `Sheet` resource for the newly created sheet.
   *
   * @example
   * ```ts
   * await sheet.addSheet({ title: 'Q3 Report', rowCount: 500 });
   * ```
   */
  async addSheet(
    props: AddSheetProps,
  ): Promise<GoogleSheetResponse<GoogleSheet>> {
    const { title, index, rowCount, columnCount } = props;
    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title,
                  index,
                  gridProperties: { rowCount, columnCount },
                },
              },
            },
          ],
        },
      });

      const added = response.data?.replies?.[0]?.addSheet;
      if (!added) {
        return {
          success: false,
          problem: `Failed to add sheet "${title}"`,
        };
      }

      return { success: true, data: added };
    } catch (error) {
      return handleError(`Error adding sheet "${title}".`, error);
    }
  }

  /**
   * Permanently deletes a sheet and all its data from the spreadsheet.
   *
   * **Warning:** this action cannot be undone programmatically. Consider
   * calling {@link duplicateSheet} first if you need a backup.
   *
   * @param sheetTitle - Exact title of the sheet to delete.
   *
   * @example
   * ```ts
   * await sheet.deleteSheet('Old Data');
   * ```
   */
  async deleteSheet(sheetTitle: string): Promise<GoogleSheetResponse<void>> {
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { requests: [{ deleteSheet: { sheetId } }] },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(`Error deleting sheet "${sheetTitle}".`, error);
    }
  }

  /**
   * Renames a sheet tab. The `fields: 'title'` mask ensures that only the
   * title property is updated and no other sheet properties are inadvertently
   * cleared.
   *
   * @param currentTitle - Current exact title of the sheet.
   * @param newTitle     - New title to apply.
   *
   * @example
   * ```ts
   * await sheet.renameSheet('Sheet1', 'Customers');
   * ```
   */
  async renameSheet(
    currentTitle: string,
    newTitle: string,
  ): Promise<GoogleSheetResponse<void>> {
    const idRes = await this._getSheetId(currentTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId, title: newTitle },
                fields: 'title',
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error renaming sheet "${currentTitle}" to "${newTitle}".`,
        error,
      );
    }
  }

  /**
   * Hides a sheet tab so it is no longer visible in the UI.
   * The sheet and all its data remain intact and can be restored with
   * {@link showSheet}.
   *
   * @param sheetTitle - Title of the sheet to hide.
   *
   * @example
   * ```ts
   * await sheet.hideSheet('Internal Notes');
   * ```
   */
  async hideSheet(sheetTitle: string): Promise<GoogleSheetResponse<void>> {
    return this._setSheetVisibility(sheetTitle, true);
  }

  /**
   * Makes a previously hidden sheet visible again.
   *
   * @param sheetTitle - Title of the sheet to show.
   *
   * @example
   * ```ts
   * await sheet.showSheet('Internal Notes');
   * ```
   */
  async showSheet(sheetTitle: string): Promise<GoogleSheetResponse<void>> {
    return this._setSheetVisibility(sheetTitle, false);
  }

  /**
   * Moves a sheet to a new position in the tab bar.
   * Indices are 0-based (0 = leftmost tab).
   *
   * @param sheetTitle - Title of the sheet to move.
   * @param newIndex   - Target 0-based position.
   *
   * @example
   * ```ts
   * // Move 'Summary' to be the first tab
   * await sheet.moveSheet('Summary', 0);
   * ```
   */
  async moveSheet(
    sheetTitle: string,
    newIndex: number,
  ): Promise<GoogleSheetResponse<void>> {
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId, index: newIndex },
                fields: 'index',
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error moving sheet "${sheetTitle}" to index ${newIndex}.`,
        error,
      );
    }
  }

  /**
   * Adds a protection rule to an entire sheet, optionally restricting
   * edits to a specific list of users.
   *
   * When `warningOnly` is `true` (the default is `false`), editors see a
   * warning dialog before editing but are not actually blocked. Set it to
   * `false` and supply `editorEmails` to enforce strict access control.
   *
   * @param props - {@link ProtectSheetProps}
   * @param props.sheetTitle   - Title of the sheet to protect.
   * @param props.description  - Optional human-readable note about why the sheet is protected.
   * @param props.warningOnly  - Show a warning instead of blocking edits. Default: `false`.
   * @param props.editorEmails - Email addresses allowed to edit despite the protection.
   * @returns The created `ProtectedRange` resource including its ID.
   *
   * @example
   * ```ts
   * await sheet.protectSheet({
   *   sheetTitle: 'Master Config',
   *   description: 'Do not edit without approval',
   *   editorEmails: ['admin@example.com'],
   * });
   * ```
   */
  async protectSheet(
    props: ProtectSheetProps,
  ): Promise<GoogleSheetResponse<ProtectedRange>> {
    const { sheetTitle, description, warningOnly, editorEmails } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              addProtectedRange: {
                protectedRange: {
                  range: { sheetId },
                  description,
                  warningOnly: warningOnly ?? false,
                  editors: editorEmails
                    ? { users: editorEmails }
                    : undefined,
                },
              },
            },
          ],
        },
      });

      const protectedRange =
        response.data?.replies?.[0]?.addProtectedRange?.protectedRange;
      if (!protectedRange) {
        return {
          success: false,
          problem: `Failed to protect sheet "${sheetTitle}"`,
        };
      }

      return { success: true, data: protectedRange };
    } catch (error) {
      return handleError(`Error protecting sheet "${sheetTitle}".`, error);
    }
  }

  // -------------------------------------------------------------------------
  // Formatting & styling
  // -------------------------------------------------------------------------

  /**
   * Sets the pixel width of one or more columns.
   * Indices are 0-based and the range is half-open `[startColumnIndex, endColumnIndex)`.
   *
   * @param props - {@link SetColumnWidthProps}
   * @param props.sheetTitle        - Title of the target sheet.
   * @param props.startColumnIndex  - 0-based first column to resize.
   * @param props.endColumnIndex    - 0-based exclusive end column.
   * @param props.pixelSize         - Desired column width in pixels.
   *
   * @example
   * ```ts
   * // Set columns A–C to 150 px wide
   * await sheet.setColumnWidth({ sheetTitle: 'Sheet1', startColumnIndex: 0, endColumnIndex: 3, pixelSize: 150 });
   * ```
   */
  async setColumnWidth(
    props: SetColumnWidthProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, startColumnIndex, endColumnIndex, pixelSize } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    return this._setDimensionPixelSize(
      idRes.data,
      'COLUMNS',
      startColumnIndex,
      endColumnIndex,
      pixelSize,
    );
  }

  /**
   * Sets the pixel height of one or more rows.
   * Indices are 0-based and the range is half-open `[startRowIndex, endRowIndex)`.
   *
   * @param props - {@link SetRowHeightProps}
   * @param props.sheetTitle     - Title of the target sheet.
   * @param props.startRowIndex  - 0-based first row to resize.
   * @param props.endRowIndex    - 0-based exclusive end row.
   * @param props.pixelSize      - Desired row height in pixels.
   *
   * @example
   * ```ts
   * // Set row 1 (the header) to 40 px tall
   * await sheet.setRowHeight({ sheetTitle: 'Sheet1', startRowIndex: 0, endRowIndex: 1, pixelSize: 40 });
   * ```
   */
  async setRowHeight(
    props: SetRowHeightProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, startRowIndex, endRowIndex, pixelSize } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    return this._setDimensionPixelSize(
      idRes.data,
      'ROWS',
      startRowIndex,
      endRowIndex,
      pixelSize,
    );
  }

  /**
   * Instructs Google Sheets to automatically fit column widths to their
   * content, equivalent to double-clicking a column border in the UI.
   *
   * @param props - {@link AutoResizeColumnsProps}
   * @param props.sheetTitle        - Title of the target sheet.
   * @param props.startColumnIndex  - 0-based first column to auto-resize.
   * @param props.endColumnIndex    - 0-based exclusive end column.
   *
   * @example
   * ```ts
   * // Auto-resize all columns A–Z (indices 0–26)
   * await sheet.autoResizeColumns({ sheetTitle: 'Report', startColumnIndex: 0, endColumnIndex: 26 });
   * ```
   */
  async autoResizeColumns(
    props: AutoResizeColumnsProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, startColumnIndex, endColumnIndex } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              autoResizeDimensions: {
                dimensions: {
                  sheetId,
                  dimension: 'COLUMNS',
                  startIndex: startColumnIndex,
                  endIndex: endColumnIndex,
                },
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error auto-resizing columns in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Freezes a number of rows and/or columns so they remain visible when
   * scrolling. Pass `0` for either count to unfreeze that dimension.
   *
   * Uses the `fields` mask `'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'`
   * to avoid clobbering other grid properties (row/column counts, etc.).
   *
   * @param props - {@link FreezeProps}
   * @param props.sheetTitle        - Title of the target sheet.
   * @param props.frozenRowCount    - Number of rows to freeze from the top. Default: `0`.
   * @param props.frozenColumnCount - Number of columns to freeze from the left. Default: `0`.
   *
   * @example
   * ```ts
   * // Freeze the header row and the first column
   * await sheet.freezeRowsAndColumns({ sheetTitle: 'Data', frozenRowCount: 1, frozenColumnCount: 1 });
   * ```
   */
  async freezeRowsAndColumns(
    props: FreezeProps,
  ): Promise<GoogleSheetResponse<void>> {
    const {
      sheetTitle,
      frozenRowCount = 0,
      frozenColumnCount = 0,
    } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: {
                  sheetId,
                  gridProperties: { frozenRowCount, frozenColumnCount },
                },
                fields:
                  'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error freezing rows/columns in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Merges a rectangular block of cells into a single cell.
   *
   * - `'MERGE_ALL'` (default) — merges the entire range into one cell.
   * - `'MERGE_COLUMNS'` — merges cells within each column independently.
   * - `'MERGE_ROWS'` — merges cells within each row independently.
   *
   * All indices are 0-based; end indices are exclusive.
   *
   * @param props - {@link MergeCellsProps}
   *
   * @example
   * ```ts
   * // Merge A1:C1 into a single header cell
   * await sheet.mergeCells({
   *   sheetTitle: 'Report',
   *   startRowIndex: 0, endRowIndex: 1,
   *   startColumnIndex: 0, endColumnIndex: 3,
   * });
   * ```
   */
  async mergeCells(
    props: MergeCellsProps,
  ): Promise<GoogleSheetResponse<void>> {
    const {
      sheetTitle,
      startRowIndex,
      endRowIndex,
      startColumnIndex,
      endColumnIndex,
      mergeType = 'MERGE_ALL',
    } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              mergeCells: {
                range: {
                  sheetId,
                  startRowIndex,
                  endRowIndex,
                  startColumnIndex,
                  endColumnIndex,
                },
                mergeType,
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error merging cells in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Applies a `CellFormat` object to every cell in a given range using a
   * `repeatCell` batchUpdate request.
   *
   * The `fields` parameter is a dot-notation mask that tells the API which
   * sub-properties of `userEnteredFormat` to update. Using `'*'` (the
   * default) replaces the entire format; use a narrower mask like
   * `'backgroundColor'` to touch only one property.
   *
   * @param props - {@link CellFormatProps}
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.range      - A1 notation range to format.
   * @param props.format     - `CellFormat` object describing the desired appearance.
   * @param props.fields     - Field mask for `userEnteredFormat`. Defaults to `'*'`.
   *
   * @example
   * ```ts
   * // Bold the header row A1:Z1
   * await sheet.setCellFormat({
   *   sheetTitle: 'Report',
   *   range: 'A1:Z1',
   *   format: { textFormat: { bold: true } },
   *   fields: 'textFormat.bold',
   * });
   * ```
   */
  async setCellFormat(
    props: CellFormatProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, range, format, fields = '*' } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              repeatCell: {
                range: this._a1ToGridRange(sheetId, range),
                cell: { userEnteredFormat: format },
                // Wrap the caller's mask inside `userEnteredFormat(...)` as
                // required by the repeatCell request schema.
                fields: `userEnteredFormat(${fields})`,
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error setting cell format in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Data utilities
  // -------------------------------------------------------------------------

  /**
   * Returns the values from the first row of a sheet — typically the column
   * headers. Delegates to {@link getRow} with `rowNumber: 1`.
   *
   * @param sheetTitle - Title of the target sheet.
   * @returns Array of header strings in column order.
   *
   * @example
   * ```ts
   * const headers = await sheet.getHeaderRow('Customers');
   * if (headers.success) console.log(headers.data); // ['Name', 'Email', 'Plan']
   * ```
   */
  async getHeaderRow(
    sheetTitle: string,
  ): Promise<GoogleSheetResponse<string[]>> {
    return this.getRow({ sheetTitle, rowNumber: 1 });
  }

  /**
   * Writes a single value into one cell identified by column letter and
   * 1-based row number. Delegates to {@link updateRange} with a 1×1 matrix.
   *
   * @param props - {@link UpdateCellProps}
   * @param props.sheetTitle       - Title of the target sheet.
   * @param props.row              - 1-based row number.
   * @param props.column           - Column letter(s), e.g. `'B'` or `'AA'`.
   * @param props.value            - Value to write.
   * @param props.valueInputOption - `'USER_ENTERED'` (default) parses dates/formulas.
   * @returns The `UpdateValuesResponse` from the API.
   *
   * @example
   * ```ts
   * await sheet.updateCell({ sheetTitle: 'Tasks', row: 5, column: 'C', value: 'DONE' });
   * ```
   */
  async updateCell(
    props: UpdateCellProps,
  ): Promise<GoogleSheetResponse<sheets_v4.Schema$UpdateValuesResponse>> {
    const { sheetTitle, row, column, value, valueInputOption } = props;
    return this.updateRange({
      sheetTitle,
      range: `${column}${row}`,
      data: [[value]],
      valueInputOption,
    });
  }

  /**
   * Reads a range and maps each data row to a typed object using the values
   * in `headerRow` as keys. This is the most ergonomic way to consume
   * structured tabular data — no manual index bookkeeping required.
   *
   * The generic parameter `T` lets callers assert the shape of the returned
   * objects; validation is not performed, so ensure the sheet structure
   * matches your type.
   *
   * @param props - {@link GetDataAsObjectsProps}
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.range      - A1 range to read. Defaults to `'A:Z'`.
   * @param props.headerRow  - 1-based row number that contains column headers. Defaults to `1`.
   * @returns Array of objects, one per data row below the header.
   *
   * @example
   * ```ts
   * type Order = { id: string; customer: string; amount: string };
   * const res = await sheet.getDataAsObjects<Order>({ sheetTitle: 'Orders' });
   * if (res.success) res.data.forEach(o => console.log(o.customer, o.amount));
   * ```
   */
  async getDataAsObjects<T = Record<string, unknown>>(
    props: GetDataAsObjectsProps,
  ): Promise<GoogleSheetResponse<T[]>> {
    const { sheetTitle, range, headerRow = 1 } = props;

    try {
      const response = await this.google_sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: `${sheetTitle}!${range ?? 'A:Z'}`,
        valueRenderOption: 'UNFORMATTED_VALUE',
      });

      const values = response.data.values ?? [];

      // If the sheet has fewer rows than the declared header row, return empty.
      if (values.length < headerRow) {
        return { success: true, data: [] };
      }

      const headers = (values[headerRow - 1] as string[]) ?? [];
      // All rows after the header row become data objects.
      const rows = values.slice(headerRow);

      const objects = rows.map((row) => {
        // Zip the header array with the row array. Missing trailing cells
        // default to an empty string to ensure every key is present.
        return Object.fromEntries(
          headers.map((header, i) => [header, row[i] ?? '']),
        ) as T;
      });

      return { success: true, data: objects };
    } catch (error) {
      return handleError(
        `Error getting data as objects from sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Copies a range of cells to a destination range using the Sheets API
   * `copyPaste` request. Both source and destination are specified in
   * A1 notation and may live on different sheets within the same spreadsheet.
   *
   * @param props - {@link CopyPasteRangeProps}
   * @param props.sheetTitle             - Source sheet title.
   * @param props.sourceRange            - A1 range to copy from.
   * @param props.destinationRange       - A1 range to paste into.
   * @param props.destinationSheetTitle  - Destination sheet title; defaults to the source sheet.
   * @param props.pasteType              - What to paste: `'PASTE_NORMAL'` (default), `'PASTE_VALUES'`, `'PASTE_FORMAT'`, or `'PASTE_FORMULA'`.
   *
   * @example
   * ```ts
   * // Copy A1:D10 values-only to the same range on 'Archive'
   * await sheet.copyPasteRange({
   *   sheetTitle: 'Live',
   *   sourceRange: 'A1:D10',
   *   destinationRange: 'A1:D10',
   *   destinationSheetTitle: 'Archive',
   *   pasteType: 'PASTE_VALUES',
   * });
   * ```
   */
  async copyPasteRange(
    props: CopyPasteRangeProps,
  ): Promise<GoogleSheetResponse<void>> {
    const {
      sheetTitle,
      sourceRange,
      destinationRange,
      destinationSheetTitle,
      pasteType = 'PASTE_NORMAL',
    } = props;

    const srcIdRes = await this._getSheetId(sheetTitle);
    if (!srcIdRes.success) return srcIdRes;

    // Only perform a second tab lookup when the destination sheet differs
    // from the source sheet — avoids an unnecessary API call otherwise.
    let dstSheetId = srcIdRes.data;
    if (destinationSheetTitle && destinationSheetTitle !== sheetTitle) {
      const dstIdRes = await this._getSheetId(destinationSheetTitle);
      if (!dstIdRes.success) return dstIdRes;
      dstSheetId = dstIdRes.data;
    }

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              copyPaste: {
                source: this._a1ToGridRange(srcIdRes.data, sourceRange),
                destination: this._a1ToGridRange(dstSheetId, destinationRange),
                pasteType,
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error copying range in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Filtering & sorting
  // -------------------------------------------------------------------------

  /**
   * Sorts rows within a range in-place according to one or more column
   * sort specifications. Later specs act as tie-breakers for earlier ones.
   *
   * @param props - {@link SortRangeProps}
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.range      - A1 range to sort, e.g. `'A2:F100'` (exclude header).
   * @param props.sortSpecs  - Array of `{ columnIndex, ascending }` objects. `columnIndex` is 0-based within the range.
   *
   * @example
   * ```ts
   * // Sort by column B ascending, then column C descending
   * await sheet.sortRange({
   *   sheetTitle: 'Sales',
   *   range: 'A2:E500',
   *   sortSpecs: [{ columnIndex: 1 }, { columnIndex: 2, ascending: false }],
   * });
   * ```
   */
  async sortRange(
    props: SortRangeProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, range, sortSpecs } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              sortRange: {
                range: this._a1ToGridRange(sheetId, range),
                sortSpecs: sortSpecs.map(
                  ({ columnIndex, ascending = true }) => ({
                    dimensionIndex: columnIndex,
                    sortOrder: ascending ? 'ASCENDING' : 'DESCENDING',
                  }),
                ),
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error sorting range in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Attaches a basic (dropdown) filter to a range, enabling users to
   * filter and hide rows directly in the Google Sheets UI.
   *
   * Only one basic filter can exist per sheet at a time. Call
   * {@link clearBasicFilter} before calling this method if a filter
   * is already active.
   *
   * @param props - {@link SetBasicFilterProps}
   * @param props.sheetTitle - Title of the target sheet.
   * @param props.range      - A1 range the filter should cover, e.g. `'A1:F200'`.
   *
   * @example
   * ```ts
   * await sheet.setBasicFilter({ sheetTitle: 'Inventory', range: 'A1:G500' });
   * ```
   */
  async setBasicFilter(
    props: SetBasicFilterProps,
  ): Promise<GoogleSheetResponse<void>> {
    const { sheetTitle, range } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              setBasicFilter: {
                filter: { range: this._a1ToGridRange(sheetId, range) },
              },
            },
          ],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error setting basic filter in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  /**
   * Removes the basic filter from a sheet, restoring all previously hidden
   * rows to visibility. Does not affect data or formatting.
   *
   * @param sheetTitle - Title of the sheet whose filter should be cleared.
   *
   * @example
   * ```ts
   * await sheet.clearBasicFilter('Inventory');
   * ```
   */
  async clearBasicFilter(
    sheetTitle: string,
  ): Promise<GoogleSheetResponse<void>> {
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [{ clearBasicFilter: { sheetId } }],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error clearing basic filter in sheet "${sheetTitle}".`,
        error,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Named ranges
  // -------------------------------------------------------------------------

  /**
   * Returns all named ranges defined in the spreadsheet.
   *
   * Named ranges provide stable, human-readable references to cell regions
   * that survive structural edits (inserted rows/columns shift the range
   * automatically). Use their IDs with {@link deleteNamedRange}.
   *
   * @returns Array of `NamedRange` objects; empty array if none exist.
   *
   * @example
   * ```ts
   * const res = await sheet.listNamedRanges();
   * if (res.success) res.data.forEach(nr => console.log(nr.name, nr.namedRangeId));
   * ```
   */
  async listNamedRanges(): Promise<GoogleSheetResponse<NamedRange[]>> {
    try {
      const response = await this.google_sheets.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
        fields: 'namedRanges',
      });
      return { success: true, data: response.data.namedRanges ?? [] };
    } catch (error) {
      return handleError('Error listing named ranges.', error);
    }
  }

  /**
   * Creates a named range within a sheet. The name must be unique across the
   * entire spreadsheet and must conform to Google Sheets naming rules
   * (no spaces, no special characters other than underscores).
   *
   * @param props - {@link AddNamedRangeProps}
   * @param props.name       - Unique name for the range (e.g. `'PRODUCT_LIST'`).
   * @param props.sheetTitle - Title of the sheet the range belongs to.
   * @param props.range      - A1 notation range to name, e.g. `'A2:A200'`.
   * @returns The created `NamedRange` resource, including its generated ID.
   *
   * @example
   * ```ts
   * await sheet.addNamedRange({
   *   name: 'TAX_RATES',
   *   sheetTitle: 'Config',
   *   range: 'B2:B10',
   * });
   * ```
   */
  async addNamedRange(
    props: AddNamedRangeProps,
  ): Promise<GoogleSheetResponse<NamedRange>> {
    const { name, sheetTitle, range } = props;
    const idRes = await this._getSheetId(sheetTitle);
    if (!idRes.success) return idRes;
    const sheetId = idRes.data;

    try {
      const response = await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [
            {
              addNamedRange: {
                namedRange: {
                  name,
                  range: this._a1ToGridRange(sheetId, range),
                },
              },
            },
          ],
        },
      });

      const namedRange =
        response.data?.replies?.[0]?.addNamedRange?.namedRange;
      if (!namedRange) {
        return {
          success: false,
          problem: `Failed to add named range "${name}"`,
        };
      }

      return { success: true, data: namedRange };
    } catch (error) {
      return handleError(`Error adding named range "${name}".`, error);
    }
  }

  /**
   * Deletes a named range by its API-assigned ID (not its human-readable name).
   * Use {@link listNamedRanges} to look up the `namedRangeId` first.
   *
   * Deleting a named range does **not** delete the underlying cells — it only
   * removes the symbolic label.
   *
   * @param namedRangeId - The `namedRangeId` string returned by {@link listNamedRanges} or {@link addNamedRange}.
   *
   * @example
   * ```ts
   * const list = await sheet.listNamedRanges();
   * const id = list.data?.find(nr => nr.name === 'OLD_RANGE')?.namedRangeId;
   * if (id) await sheet.deleteNamedRange(id);
   * ```
   */
  async deleteNamedRange(
    namedRangeId: string,
  ): Promise<GoogleSheetResponse<void>> {
    try {
      await this.google_sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [{ deleteNamedRange: { namedRangeId } }],
        },
      });
      return { success: true, data: undefined };
    } catch (error) {
      return handleError(
        `Error deleting named range "${namedRangeId}".`,
        error,
      );
    }
  }
}
