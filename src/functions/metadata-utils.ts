// Copyright IBM Corp. 2025

/**
 * Shared utilities and metadata for API operations
 * This module contains common configurations and utilities used across different API handlers
 */

import {
  Location,
  Mobile,
  Fugitive,
  Stationary,
  Calculation,
  TransportationAndDistribution,
} from "emissions-api-sdk";

/**
 * Valid API names supported by the system
 */
export const VALID_API_NAMES = [
  "location",
  "mobile",
  "fugitive",
  "stationary",
  "calculation",
  "transportationanddistribution",
  "factor",
] as const;

/**
 * API class mapping for metadata operations
 * Maps lowercase API names to their SDK classes
 */
export const API_CLASS_MAP = {
  location: Location,
  mobile: Mobile,
  fugitive: Fugitive,
  stationary: Stationary,
  calculation: Calculation,
  transportationanddistribution: TransportationAndDistribution,
} as const;

/**
 * Configuration for API types operations
 */
export interface ApiTypesConfig {
  name: string;
  getTypes: () => Promise<any>;
}

/**
 * Configuration for API area operations
 */
export interface ApiAreaConfig {
  name: string;
  class: typeof Location | typeof Mobile | typeof Fugitive | typeof Stationary | typeof Calculation | typeof TransportationAndDistribution;
}

/**
 * All available API configurations for types operations
 * Column A: Location, B: Mobile, C: Fugitive, D: Stationary, E: Calculation, F: TransportationAndDistribution
 */
export const API_TYPES_CONFIGS: ApiTypesConfig[] = [
  { name: "Location", getTypes: Location.getTypes },
  { name: "Mobile", getTypes: Mobile.getTypes },
  { name: "Fugitive", getTypes: Fugitive.getTypes },
  { name: "Stationary", getTypes: Stationary.getTypes },
  { name: "Calculation", getTypes: Calculation.getTypes },
  { name: "TransportationAndDistribution", getTypes: TransportationAndDistribution.getTypes },
];

/**
 * All available API configurations for area operations
 * Only stores data for 2 representative APIs:
 * - calculation (represents: calculation, location, factor)
 * - mobile (represents: mobile, stationary, fugitive)
 */
export const API_AREA_CONFIGS: ApiAreaConfig[] = [
  { name: "calculation", class: Calculation },
  { name: "mobile", class: Mobile },
];

/**
 * Maps API names to their representative API for area data
 * Group 1 (calculation): calculation, location, factor
 * Group 2 (mobile): mobile, stationary, fugitive
 */
export const API_AREA_MAPPING: Record<string, string> = {
  calculation: "calculation",
  location: "calculation",
  factor: "calculation",
  mobile: "mobile",
  stationary: "mobile",
  fugitive: "mobile",
  transportationanddistribution: "mobile", // Maps to mobile group
};

/**
 * Configuration for data refresh mechanism
 * For testing: Set REFRESH_INTERVAL_MS to a short duration (e.g., 10000 = 10 seconds)
 * For production: Use 2 days (2 * 24 * 60 * 60 * 1000)
 */
export const REFRESH_CONFIG = {
  // Change this value for testing: 10 * 1000 (10 seconds) or 30 * 1000 (30 seconds)
  // Production value: 2 * 24 * 60 * 60 * 1000 (2 days)
  REFRESH_INTERVAL_MS: 2 * 24 * 60 * 60 * 1000, // 2 days
  REFRESH_INTERVAL_DAYS: 2,
};

/**
 * Sheet metadata interface
 */
export interface SheetMetadata {
  timestamp: number;
}

/**
 * Gets metadata from a hidden sheet
 * @param sheetName The name of the sheet
 * @returns Sheet metadata or null if not found
 */
export async function getSheetMetadata(sheetName: string): Promise<SheetMetadata | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const metadataRange = sheet.getRangeByIndexes(0, 0, 1, 2);
      metadataRange.load("values");
      await context.sync();

      const values = metadataRange.values[0];
      if (values[0] !== "METADATA") {
        return null; // No metadata found
      }

      return {
        timestamp: parseInt(values[1] as string),
      };
    });
  } catch (error) {
    console.error(`Error reading metadata from ${sheetName}:`, error);
    return null;
  }
}

/**
 * Sets metadata in a hidden sheet
 * @param sheetName The name of the sheet
 * @param metadata The metadata to store
 */
export async function setSheetMetadata(sheetName: string, metadata: SheetMetadata): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const metadataRange = sheet.getRangeByIndexes(0, 0, 1, 2);
    metadataRange.values = [["METADATA", metadata.timestamp.toString()]];
    await context.sync();
  });
}

/**
 * Checks if sheet data is stale based on age only
 * Data is considered stale if it's older than REFRESH_INTERVAL_MS (2 days by default)
 *
 * @param sheetName The name of the sheet to check
 * @returns True if data is stale and needs refresh
 */
export async function isSheetDataStale(sheetName: string): Promise<boolean> {
  try {
    const metadata = await getSheetMetadata(sheetName);

    // Metadata should always exist since we create sheets with it
    // If it doesn't exist, getSheetMetadata returns null and we'll refresh
    if (!metadata) {
      return true;
    }

    // Check age only
    const age = Date.now() - metadata.timestamp;
    return age > REFRESH_CONFIG.REFRESH_INTERVAL_MS;
  } catch (error) {
    console.error(`Error checking staleness for ${sheetName}:`, error);
    return true; // On error, assume stale
  }
}

/**
 * Deletes a sheet if it exists
 * @param sheetName The name of the sheet to delete
 */
export async function deleteSheetIfExists(sheetName: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheet = sheets.items.find((s) => s.name === sheetName);
      if (sheet) {
        sheet.delete();
        await context.sync();
      }
    });
  } catch (error) {
    console.error(`Error deleting sheet ${sheetName}:`, error);
  }
}

/**
 * Checks if a sheet exists
 * @param sheetName The name of the sheet to check
 * @returns True if sheet exists
 */
export async function sheetExists(sheetName: string): Promise<boolean> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      
      const sheet = sheets.items.find((s) => s.name === sheetName);
      return !!sheet;
    });
  } catch (error) {
    console.error(`Error checking if sheet ${sheetName} exists:`, error);
    return false;
  }
}

/**
 * Refreshes a sheet if it exists and is stale
 * Only refreshes if sheet already exists - does not create new sheets
 * @param sheetName The name of the sheet
 * @param recreateFunction Function to recreate the sheet with fresh data
 * @returns True if sheet was refreshed
 */
export async function refreshSheetIfStale(
  sheetName: string,
  recreateFunction: () => Promise<void>
): Promise<boolean> {
  const exists = await sheetExists(sheetName);
  
  if (!exists) {
    return false;
  }

  const isStale = await isSheetDataStale(sheetName);

  if (isStale) {
    await deleteSheetIfExists(sheetName);
    await recreateFunction();
    return true;
  }

  return false;
}

/**
 * Always refreshes a sheet if it exists (used for login trigger)
 * Only refreshes if sheet already exists - does not create new sheets
 * @param sheetName The name of the sheet
 * @param recreateFunction Function to recreate the sheet with fresh data
 * @returns True if sheet was refreshed
 */
export async function refreshSheetOnLogin(
  sheetName: string,
  recreateFunction: () => Promise<void>
): Promise<boolean> {
  const exists = await sheetExists(sheetName);
  
  if (!exists) {
    return false;
  }

  await deleteSheetIfExists(sheetName);
  await recreateFunction();
  return true;
}

/**
 * Validates if the provided API name is valid
 * @param apiName The API name to validate
 * @returns The normalized API name if valid
 * @throws CustomFunctions.Error if invalid
 */
export function validateApiName(apiName: string): string {
  const normalizedApiName = apiName.toLowerCase().trim();
  
  if (!VALID_API_NAMES.includes(normalizedApiName as any)) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `Invalid API name. Valid options: ${VALID_API_NAMES.join(", ")}`
    );
  }
  
  return normalizedApiName;
}

/**
 * Gets the target cell from a cell address
 * @param context Excel context
 * @param cellAddress The address of the cell (e.g., "Sheet1!A1" or "A1")
 * @returns The target Excel.Range
 */
export async function getTargetCell(
  context: Excel.RequestContext,
  cellAddress: string
): Promise<Excel.Range> {
  // Parse the cell address to get sheet name and cell reference
  const [sheetName, cellRef] = cellAddress.includes("!")
    ? cellAddress.split("!")
    : ["", cellAddress];
  
  // Get the target cell
  if (sheetName) {
    const targetSheet = context.workbook.worksheets.getItem(sheetName);
    return targetSheet.getRange(cellRef);
  } else {
    // If no sheet name, use active sheet
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    return activeSheet.getRange(cellRef);
  }
}

/**
 * Applies list validation to a cell
 * @param targetCell The cell to apply validation to
 * @param values Array of values for the dropdown
 * @param title Error alert title
 * @param message Error alert message
 */
export function applyListValidation(
  targetCell: Excel.Range,
  values: string[],
  title: string,
  message: string
): void {
  // Clear any existing validation
  targetCell.dataValidation.clear();
  
  // Create comma-separated list
  const valuesList = values.join(",");
  
  // Apply list validation
  targetCell.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: valuesList,
    },
  };
  
  targetCell.dataValidation.errorAlert = {
    showAlert: true,
    style: Excel.DataValidationAlertStyle.stop,
    title,
    message,
  };
}

/**
 * Handles errors in custom functions, re-throwing CustomFunctions.Error as-is
 * @param error The error to handle
 * @param defaultMessage Default error message if not a CustomFunctions.Error
 * @throws CustomFunctions.Error
 */
export function handleCustomFunctionError(error: unknown, defaultMessage: string): never {
  // Re-throw CustomFunctions.Error as-is
  if (error instanceof CustomFunctions.Error || (error as any)?.name === "CustomFunctions.Error") {
    throw error;
  }
  throw new CustomFunctions.Error(
    CustomFunctions.ErrorCode.notAvailable,
    `${defaultMessage}: ${(error as any)?.message || "Unknown error"}`
  );
}
