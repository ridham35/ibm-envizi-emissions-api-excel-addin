// Copyright IBM Corp. 2025

import { ensureClient } from "./client";
import { API_TYPES_CONFIGS, setSheetMetadata } from "./metadata-utils";

/**
 * Map of API names to their fixed column indices (0-based)
 * Note: factor uses the same column as calculation (same types)
 */
export const API_COLUMN_MAP: Record<string, number> = {
  location: 0,
  mobile: 1,
  fugitive: 2,
  stationary: 3,
  calculation: 4,
  transportationanddistribution: 5,
  factor: 4, // Factor uses same types as calculation
};

/**
 * Sheet name for storing API types
 */
const API_TYPES_SHEET_NAME = "API_Types_Data";

/**
 * Fetches types from all API endpoints
 * @returns Promise with array of API names and their types
 */
export async function fetchAllApiTypes(): Promise<Map<string, string[]>> {
  await ensureClient();

  const apiTypesMap = new Map<string, string[]>();

  // Fetch types from all APIs in parallel for better performance
  const fetchPromises = API_TYPES_CONFIGS.map(async (config) => {
    try {
      const response = await config.getTypes();
      // The SDK returns an object with a 'types' array property
      const types = response?.types || [];
      return { name: config.name, types };
    } catch (error) {
      console.error(`Error fetching types for ${config.name}:`, error);
      return { name: config.name, types: [] };
    }
  });

  const results = await Promise.all(fetchPromises);

  // Populate the map with results
  results.forEach((result) => {
    apiTypesMap.set(result.name, result.types);
  });

  return apiTypesMap;
}

/**
 * Creates or gets the API types sheet
 * @param context Excel context
 * @returns The worksheet for API types
 */
async function getOrCreateApiTypesSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  let sheet = sheets.items.find((s) => s.name === API_TYPES_SHEET_NAME);

  if (!sheet) {
    sheet = sheets.add(API_TYPES_SHEET_NAME);
    // comment the line below to unhide the sheet
    sheet.visibility = Excel.SheetVisibility.hidden;
  }

  return sheet;
}

/**
 * Writes API types to Excel sheet in fixed column positions
 * @param apiTypesMap Map of API names to their types
 */
export async function writeApiTypesToSheet(apiTypesMap: Map<string, string[]>): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrCreateApiTypesSheet(context);

    // Clear existing content
    const usedRange = sheet.getUsedRange(true);
    usedRange.clear();

    // Find the maximum number of types across all APIs
    let maxTypes = 0;
    apiTypesMap.forEach((types) => {
      maxTypes = Math.max(maxTypes, types.length);
    });

    // Fixed column count (6 APIs)
    const columnCount = API_TYPES_CONFIGS.length;
    const rowCount = maxTypes + 2; // +1 for metadata row, +1 for header row
    const data: (string | null)[][] = Array(rowCount)
      .fill(null)
      .map(() => Array(columnCount).fill(null));

    // Row 0: Metadata (will be set separately)
    // Row 1: Headers
    // Row 2+: Data

    // Fill in the data in fixed column order matching API_TYPES_CONFIGS
    API_TYPES_CONFIGS.forEach((config, columnIndex) => {
      const types = apiTypesMap.get(config.name) || [];
      
      // Set header (row 1)
      data[1][columnIndex] = config.name;

      // Set types (starting from row 2)
      types.forEach((type, typeIndex) => {
        data[typeIndex + 2][columnIndex] = type;
      });
    });

    // Write data to sheet (starting from row 1, row 0 is for metadata)
    const range = sheet.getRangeByIndexes(1, 0, rowCount - 1, columnCount);
    range.values = data.slice(1);

    // Format headers (row 1)
    const headerRange = sheet.getRangeByIndexes(1, 0, 1, columnCount);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Auto-fit columns
    range.format.autofitColumns();

    await context.sync();

    // Store metadata in row 0
    await setSheetMetadata(API_TYPES_SHEET_NAME, {
      timestamp: Date.now(),
    });

    console.log(`✅ Successfully wrote API types to sheet: ${API_TYPES_SHEET_NAME} with metadata`);
  });
}

/**
 * Main function to load and populate API types
 * This should be called after successful login
 */
export async function loadAndPopulateApiTypes(): Promise<void> {
  try {
    console.log("Fetching API types...");
    const apiTypesMap = await fetchAllApiTypes();

    console.log("Writing API types to Excel sheet...");
    await writeApiTypesToSheet(apiTypesMap);

    console.log("API types loaded successfully!");
  } catch (error) {
    console.error("Error loading API types:", error);
    throw error;
  }
}
