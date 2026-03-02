import { createOAuth2Client } from '../auth/create-oauth2-client.js';
import type { GoogleAuthConfig } from '../auth/types.js';
import { GoogleSpreadSheet } from './google-spreadsheet.js';

export interface GoogleSheetsConfig extends GoogleAuthConfig {
  refreshToken: string;
}

/**
 * Single-call factory: creates an OAuth2 client, sets the refresh token
 * credentials, and returns a ready-to-use GoogleSpreadSheet instance.
 */
export function createSheetsClient(
  config: GoogleSheetsConfig,
  spreadsheetId: string,
): GoogleSpreadSheet {
  const { refreshToken, ...authConfig } = config;
  const oauth2Client = createOAuth2Client(authConfig);
  oauth2Client.setCredentials({ refresh_token: refreshToken });
  return new GoogleSpreadSheet({ auth: oauth2Client, spreadsheetId });
}
