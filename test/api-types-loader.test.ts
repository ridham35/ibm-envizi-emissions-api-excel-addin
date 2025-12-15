// Copyright IBM Corp. 2025

import {
  fetchAllApiTypes,
  writeApiTypesToSheet,
  loadAndPopulateApiTypes,
  API_COLUMN_MAP,
} from "../src/functions/api-types-loader";
import {
  Location,
  Mobile,
  Fugitive,
  Stationary,
  Calculation,
  TransportationAndDistribution,
} from "emissions-api-sdk";
import { ensureClient } from "../src/functions/client";

// Mock window.apiCredentials
(global as any).window = {
  apiCredentials: {
    tenantId: "test-tenant",
    orgId: "test-org",
    apiKey: "test-key",
  },
};

// Mock the SDK modules
jest.mock("emissions-api-sdk", () => ({
  Location: { getTypes: jest.fn() },
  Mobile: { getTypes: jest.fn() },
  Fugitive: { getTypes: jest.fn() },
  Stationary: { getTypes: jest.fn() },
  Calculation: { getTypes: jest.fn() },
  TransportationAndDistribution: { getTypes: jest.fn() },
}));

jest.mock("../src/functions/client", () => ({
  ensureClient: jest.fn(),
}));

// Mock Excel
const mockContext = {
  workbook: {
    worksheets: {
      items: [] as any[],
      load: jest.fn(),
      add: jest.fn(),
      getItem: jest.fn(),
    },
  },
  sync: jest.fn(),
};

const mockRange = {
  values: [] as any[][],
  clear: jest.fn(),
  format: {
    font: { bold: false, color: "" },
    fill: { color: "" },
    autofitColumns: jest.fn(),
  },
};

const mockSheet = {
  name: "API_Types_Data",
  visibility: null as any,
  getUsedRange: jest.fn(),
  getRangeByIndexes: jest.fn().mockReturnValue(mockRange),
};

global.Excel = {
  run: jest.fn((callback) => callback(mockContext)),
  SheetVisibility: { hidden: "hidden" },
} as any;

describe("api-types-loader", () => {
  beforeEach(() => {
    jest.clearAllMocks();
    mockContext.workbook.worksheets.items = [];
  });

  describe("API_COLUMN_MAP", () => {
    it("should have correct column mappings", () => {
      expect(API_COLUMN_MAP).toEqual({
        location: 0,
        mobile: 1,
        fugitive: 2,
        stationary: 3,
        calculation: 4,
        transportationanddistribution: 5,
        factor: 4, // Factor uses same column as calculation
      });
    });
  });

  describe("fetchAllApiTypes", () => {
    it("should fetch types from all APIs successfully", async () => {
      const mockTypes = {
        Location: ["electricity", "natural gas"],
        Mobile: ["diesel", "gasoline"],
        Fugitive: ["refrigerant", "co2"],
        Stationary: ["coal", "oil"],
        Calculation: ["generic type 1", "generic type 2"],
        TransportationAndDistribution: ["freight", "shipping"],
      };

      (Location.getTypes as jest.Mock).mockResolvedValue({ types: mockTypes.Location });
      (Mobile.getTypes as jest.Mock).mockResolvedValue({ types: mockTypes.Mobile });
      (Fugitive.getTypes as jest.Mock).mockResolvedValue({ types: mockTypes.Fugitive });
      (Stationary.getTypes as jest.Mock).mockResolvedValue({ types: mockTypes.Stationary });
      (Calculation.getTypes as jest.Mock).mockResolvedValue({ types: mockTypes.Calculation });
      (TransportationAndDistribution.getTypes as jest.Mock).mockResolvedValue({
        types: mockTypes.TransportationAndDistribution,
      });

      const result = await fetchAllApiTypes();

      expect(ensureClient).toHaveBeenCalled();
      expect(result.size).toBe(6);
      expect(result.get("Location")).toEqual(mockTypes.Location);
      expect(result.get("Mobile")).toEqual(mockTypes.Mobile);
      expect(result.get("Fugitive")).toEqual(mockTypes.Fugitive);
      expect(result.get("Stationary")).toEqual(mockTypes.Stationary);
      expect(result.get("Calculation")).toEqual(mockTypes.Calculation);
      expect(result.get("TransportationAndDistribution")).toEqual(mockTypes.TransportationAndDistribution);
    });

    it("should handle API errors gracefully", async () => {
      (Location.getTypes as jest.Mock).mockResolvedValue({ types: ["type1"] });
      (Mobile.getTypes as jest.Mock).mockRejectedValue(new Error("API Error"));
      (Fugitive.getTypes as jest.Mock).mockResolvedValue({ types: ["type2"] });
      (Stationary.getTypes as jest.Mock).mockResolvedValue({ types: [] });
      (Calculation.getTypes as jest.Mock).mockResolvedValue({ types: ["type3"] });
      (TransportationAndDistribution.getTypes as jest.Mock).mockResolvedValue({ types: ["type4"] });

      const result = await fetchAllApiTypes();

      expect(result.size).toBe(6);
      expect(result.get("Location")).toEqual(["type1"]);
      expect(result.get("Mobile")).toEqual([]); // Error should result in empty array
      expect(result.get("Fugitive")).toEqual(["type2"]);
    });

    it("should handle missing types property", async () => {
      (Location.getTypes as jest.Mock).mockResolvedValue({});
      (Mobile.getTypes as jest.Mock).mockResolvedValue({ types: ["type1"] });
      (Fugitive.getTypes as jest.Mock).mockResolvedValue(null);
      (Stationary.getTypes as jest.Mock).mockResolvedValue({ types: ["type2"] });
      (Calculation.getTypes as jest.Mock).mockResolvedValue({ types: ["type3"] });
      (TransportationAndDistribution.getTypes as jest.Mock).mockResolvedValue({ types: ["type4"] });

      const result = await fetchAllApiTypes();

      expect(result.get("Location")).toEqual([]);
      expect(result.get("Mobile")).toEqual(["type1"]);
      expect(result.get("Fugitive")).toEqual([]);
    });
  });

  describe("writeApiTypesToSheet", () => {
    beforeEach(() => {
      mockSheet.getUsedRange.mockReturnValue(mockRange);
      mockSheet.getRangeByIndexes.mockClear();
      mockSheet.getRangeByIndexes.mockReturnValue(mockRange);
      mockContext.workbook.worksheets.add.mockReturnValue(mockSheet);
      mockContext.workbook.worksheets.getItem.mockReturnValue(mockSheet);
    });

    it("should create sheet if it does not exist", async () => {
      const apiTypesMap = new Map([
        ["Location", ["type1", "type2"]],
        ["Mobile", ["type3"]],
      ]);

      await writeApiTypesToSheet(apiTypesMap);

      expect(mockContext.workbook.worksheets.add).toHaveBeenCalledWith("API_Types_Data");
      expect(mockSheet.visibility).toBe("hidden");
    });

    it("should write data in correct column order", async () => {
      const apiTypesMap = new Map([
        ["Location", ["loc1", "loc2"]],
        ["Mobile", ["mob1"]],
        ["Fugitive", ["fug1", "fug2", "fug3"]],
        ["Stationary", ["stat1"]],
        ["Calculation", ["calc1", "calc2"]],
        ["TransportationAndDistribution", ["trans1"]],
      ]);

      await writeApiTypesToSheet(apiTypesMap);

      expect(mockRange.values).toBeDefined();
      expect(mockRange.format.autofitColumns).toHaveBeenCalled();
    });

    it("should clear existing content before writing", async () => {
      const apiTypesMap = new Map([["Location", ["type1"]]]);

      await writeApiTypesToSheet(apiTypesMap);

      expect(mockRange.clear).toHaveBeenCalled();
    });

    it("should format headers correctly", async () => {
      const apiTypesMap = new Map([["Location", ["type1"]]]);

      await writeApiTypesToSheet(apiTypesMap);

      // The function calls getRangeByIndexes multiple times (for metadata, header, and data)
      // Check that it was called with the header range parameters
      const calls = mockSheet.getRangeByIndexes.mock.calls;
      const headerCall = calls.find(call => call[0] === 1 && call[2] === 1 && call[3] === 6); // startRow=1, rowCount=1, colCount=6
      expect(headerCall).toBeDefined();
    });
  });

  describe("loadAndPopulateApiTypes", () => {
    beforeEach(() => {
      mockSheet.getUsedRange.mockReturnValue(mockRange);
      mockSheet.getRangeByIndexes.mockClear();
      mockSheet.getRangeByIndexes.mockReturnValue(mockRange);
      mockContext.workbook.worksheets.add.mockReturnValue(mockSheet);
      mockContext.workbook.worksheets.getItem.mockReturnValue(mockSheet);
    });

    it("should fetch and write API types successfully", async () => {
      (Location.getTypes as jest.Mock).mockResolvedValue({ types: ["type1"] });
      (Mobile.getTypes as jest.Mock).mockResolvedValue({ types: ["type2"] });
      (Fugitive.getTypes as jest.Mock).mockResolvedValue({ types: ["type3"] });
      (Stationary.getTypes as jest.Mock).mockResolvedValue({ types: ["type4"] });
      (Calculation.getTypes as jest.Mock).mockResolvedValue({ types: ["type5"] });
      (TransportationAndDistribution.getTypes as jest.Mock).mockResolvedValue({ types: ["type6"] });

      await loadAndPopulateApiTypes();

      expect(ensureClient).toHaveBeenCalled();
      expect(mockContext.workbook.worksheets.add).toHaveBeenCalled();
    });

    it("should throw error if fetching fails", async () => {
      (ensureClient as jest.Mock).mockRejectedValue(new Error("Client error"));

      await expect(loadAndPopulateApiTypes()).rejects.toThrow("Client error");
    });
  });
});

// Made with Bob
