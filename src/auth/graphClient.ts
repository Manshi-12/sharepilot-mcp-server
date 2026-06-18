import { ConfidentialClientApplication } from "@azure/msal-node";
import axios, { AxiosInstance } from "axios";
import dotenv from "dotenv";

dotenv.config();

const TENANT_ID = process.env.TENANT_ID || "";
const CLIENT_ID = process.env.CLIENT_ID || "";
const CLIENT_SECRET = process.env.CLIENT_SECRET || "";
const SITE_URL = process.env.SITE_URL || "";

let cca: ConfidentialClientApplication | null = null;

// ── Token cache ───────────────────────────────────────────────────────────────
// MSAL's acquireTokenByClientCredential already caches internally, but we keep
// a thin wrapper cache so we don't even call MSAL on every request.
// We refresh 2 minutes before expiry to avoid edge-case clock skew.
interface TokenCache {
  token: string;
  expiresAt: number; // epoch ms
}
let tokenCache: TokenCache | null = null;
const TOKEN_REFRESH_BUFFER_MS = 2 * 60 * 1000; // 2 minutes

function getConfidentialClient(): ConfidentialClientApplication {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error(
      "Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET environment variables. " +
      "Set these in .env (local) or Application Settings (Azure App Service)."
    );
  }

  if (!cca) {
    cca = new ConfidentialClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET,
      },
    });
  }

  return cca;
}

export async function getAccessToken(): Promise<string> {
  // Return cached token if still valid
  if (tokenCache && tokenCache.expiresAt - TOKEN_REFRESH_BUFFER_MS > Date.now()) {
    return tokenCache.token;
  }

  const client = getConfidentialClient();

  const result = await client.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  if (!result || !result.accessToken) {
    throw new Error("Failed to acquire access token from Microsoft Graph.");
  }

  // MSAL returns expiresOn as a Date — convert to epoch ms
  const expiresAt = result.expiresOn
    ? result.expiresOn.getTime()
    : Date.now() + 55 * 60 * 1000; // fallback: 55 min

  tokenCache = { token: result.accessToken, expiresAt };

  return result.accessToken;
}

export async function getGraphClient(): Promise<AxiosInstance> {
  const token = await getAccessToken();

  return axios.create({
    baseURL: "https://graph.microsoft.com/v1.0",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    // Fix #11 — timeout so Graph hangs don't hang the server forever
    timeout: 30000, // 30 seconds
  });
}

// ── SharePoint REST (_api) token ────────────────────────────────────────────
// IMPORTANT: this is a DIFFERENT resource/audience than Microsoft Graph.
// The legacy SharePoint REST endpoints (/_api/...) reject Graph-scoped tokens
// with a 401. We need a token whose audience is the SharePoint site itself,
// acquired with scope "<SITE_URL>/.default". Cached separately from the
// Graph token since they are not interchangeable and expire independently.
let spTokenCache: TokenCache | null = null;

export async function getSharePointAccessToken(): Promise<string> {
  if (!SITE_URL) {
    throw new Error("Missing SITE_URL environment variable — required to acquire a SharePoint REST API token.");
  }

  if (spTokenCache && spTokenCache.expiresAt - TOKEN_REFRESH_BUFFER_MS > Date.now()) {
    return spTokenCache.token;
  }

  const client = getConfidentialClient();
  const siteOrigin = new URL(SITE_URL).origin; // e.g. https://dwivedimanshi12outlook.sharepoint.com

  const result = await client.acquireTokenByClientCredential({
    scopes: [`${siteOrigin}/.default`],
  });

  if (!result || !result.accessToken) {
    throw new Error("Failed to acquire access token for the SharePoint REST API.");
  }

  const expiresAt = result.expiresOn
    ? result.expiresOn.getTime()
    : Date.now() + 55 * 60 * 1000;

  spTokenCache = { token: result.accessToken, expiresAt };

  return result.accessToken;
}