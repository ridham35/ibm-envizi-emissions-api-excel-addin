// Copyright IBM Corp. 2025

import { Client, ClientConfig } from "emissions-api-sdk";

import { getApiUrl } from "../common/env";
import {
  ApiCredentials,
  getApiCredentials,
  loadApiCredentialsFromStorage,
} from "../common/credentials";

async function getClientConfig(apiCredentials?: ApiCredentials): Promise<ClientConfig> {
  let resolvedApiCredentials = apiCredentials || getApiCredentials();
  if (!resolvedApiCredentials) {
    resolvedApiCredentials = await loadApiCredentialsFromStorage();
    if (!resolvedApiCredentials) {
      Office.addin.showAsTaskpane();
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "Enter your credentials in the task pane."
      );
    }
  }
  const config: ClientConfig = {
    host: getApiUrl("ghgemissions"),
    authUrl: `${getApiUrl("saascore")}/authentication-retrieve/api-key`,
    apiKey: resolvedApiCredentials.apiKey,
    clientId: resolvedApiCredentials.tenantId,
    orgId: resolvedApiCredentials.orgId,
    isExcelAddIn: true,
  };
  return config;
}

/**
 * Ensures the Client object is properly initialized.
 */
export async function ensureClient(apiCredentials?: ApiCredentials): Promise<void> {
  if (Client["instance"] && !apiCredentials) {
    return;
  }
  const config = await getClientConfig(apiCredentials);
  return Client.getClient(config);
}

/**
 * Resets the client instance.
 */
export function resetClient(): void {
  // Need to find a better way to reset later.
  Client["instance"] = null;
}
