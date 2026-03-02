import { google } from 'googleapis';
import type { Auth } from 'googleapis';
import type { GoogleAuthConfig } from './types.js';

const DEFAULT_SCOPES = [
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/spreadsheets',
];

/**
 * Create a reusable OAuth2 client from plain config.
 */
export function createOAuth2Client(config: GoogleAuthConfig): Auth.OAuth2Client {
  const { clientId, clientSecret, redirectUri } = config;
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

/**
 * Generate the consent-screen URL. Scopes default to sheets + drive + gmail.send.
 */
export function generateAuthUrl(
  client: Auth.OAuth2Client,
  scopes: string[] = DEFAULT_SCOPES,
): string {
  return client.generateAuthUrl({
    access_type: 'offline',
    scope: scopes,
  });
}

/**
 * Exchange an auth code for a refresh token. Call once to bootstrap credentials.
 */
export async function retrieveRefreshToken(
  client: Auth.OAuth2Client,
  code: string,
): Promise<string | null> {
  const { tokens } = await client.getToken(code);
  return tokens.refresh_token ?? null;
}
