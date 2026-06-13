import { ConfidentialClientApplication } from "@azure/msal-node";
import axios, { AxiosInstance } from "axios";
import dotenv from "dotenv";

dotenv.config();

const TENANT_ID = process.env.TENANT_ID || "";
const CLIENT_ID = process.env.CLIENT_ID || "";
const CLIENT_SECRET = process.env.CLIENT_SECRET || "";

let cca: ConfidentialClientApplication | null = null;

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
  const client = getConfidentialClient();

  const result = await client.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  if (!result || !result.accessToken) {
    throw new Error("Failed to acquire access token from Microsoft Graph.");
  }

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
  });
}