// Copyright IBM Corp. 2025

import {
  validateApiName,
  VALID_API_NAMES,
  API_TYPES_CONFIGS,
  API_AREA_CONFIGS,
  REFRESH_CONFIG,
  SheetMetadata,
  getSheetMetadata,
  setSheetMetadata,
  isSheetDataStale,
  sheetExists,
  deleteSheetIfExists,
  refreshSheetIfStale,
  refreshSheetOnLogin,
} from "../src/functions/metadata-utils";

// Mock CustomFunctions
class CustomFunctionsError extends Error {
  code: string;
  constructor(code: string, message: string) {
    super(message);
    this.code = code;
    this.name = "CustomFunctions.Error";
  }
}

global.CustomFunctions = {
  Error: CustomFunctionsError,
  ErrorCode: {
    invalidValue: "InvalidValue",
    notAvailable: "NotAvailable",
  },
} as any;

describe("metadata-utils", () => {
  describe("VALID_API_NAMES", () => {
    it("should export all valid API names", () => {
      expect(VALID_API_NAMES).toEqual([
        "location",
        "mobile",
        "fugitive",
        "stationary",
        "calculation",
        "transportationanddistribution",
        "factor",
      ]);
    });
  });

  describe("API_TYPES_CONFIGS", () => {
    it("should have 6 API configurations", () => {
      expect(API_TYPES_CONFIGS).toHaveLength(6);
    });

    it("should have correct API names", () => {
      const names = API_TYPES_CONFIGS.map((config) => config.name);
      expect(names).toEqual([
        "Location",
        "Mobile",
        "Fugitive",
        "Stationary",
        "Calculation",
        "TransportationAndDistribution",
      ]);
    });

    it("should have getTypes method for each config", () => {
      API_TYPES_CONFIGS.forEach((config) => {
        expect(config.getTypes).toBeDefined();
        expect(typeof config.getTypes).toBe("function");
      });
    });
  });

  describe("API_AREA_CONFIGS", () => {
    it("should have 2 representative API configurations (optimization for area data)", () => {
      expect(API_AREA_CONFIGS).toHaveLength(2);
    });

    it("should have correct representative API names in lowercase", () => {
      const names = API_AREA_CONFIGS.map((config) => config.name);
      expect(names).toEqual([
        "calculation",
        "mobile",
      ]);
    });

    it("should have class property for each config", () => {
      API_AREA_CONFIGS.forEach((config) => {
        expect(config.class).toBeDefined();
        expect(typeof config.class).toBe("object");
      });
    });

    it("should NOT include factor or factorsearch", () => {
      const names = API_AREA_CONFIGS.map((config) => config.name);
      expect(names).not.toContain("factor");
      expect(names).not.toContain("factorsearch");
    });
  });

  describe("validateApiName", () => {
    it("should accept valid API names", () => {
      VALID_API_NAMES.forEach((apiName) => {
        expect(validateApiName(apiName)).toBe(apiName);
      });
    });

    it("should normalize API names to lowercase", () => {
      expect(validateApiName("LOCATION")).toBe("location");
      expect(validateApiName("Mobile")).toBe("mobile");
      expect(validateApiName("FuGiTiVe")).toBe("fugitive");
    });

    it("should trim whitespace", () => {
      expect(validateApiName("  location  ")).toBe("location");
      expect(validateApiName("\tmobile\n")).toBe("mobile");
    });

    it("should throw error for invalid API names", () => {
      expect(() => validateApiName("invalid")).toThrow();
      expect(() => validateApiName("unknown")).toThrow();
      expect(() => validateApiName("")).toThrow();
      
      // Verify it's a CustomFunctions.Error
      try {
        validateApiName("invalid");
        fail("Should have thrown error");
      } catch (error: any) {
        expect(error.code).toBe("InvalidValue");
      }
    });

    it("should include valid options in error message", () => {
      try {
        validateApiName("invalid");
        fail("Should have thrown error");
      } catch (error: any) {
        expect(error.message).toContain("Invalid API name");
        expect(error.message).toContain("location");
        expect(error.message).toContain("mobile");
        expect(error.code).toBe("InvalidValue");
      }
    });
  });

  describe("REFRESH_CONFIG", () => {
    it("should have correct refresh interval", () => {
      expect(REFRESH_CONFIG.REFRESH_INTERVAL_MS).toBe(2 * 24 * 60 * 60 * 1000); // 2 days
      expect(REFRESH_CONFIG.REFRESH_INTERVAL_DAYS).toBe(2);
    });
  });

  describe("Refresh Mechanism", () => {
    // Mock Excel context
    const mockContext = {
      workbook: {
        worksheets: {
          items: [] as any[],
          load: jest.fn(),
          getItem: jest.fn(),
        },
      },
      sync: jest.fn().mockResolvedValue(undefined),
    };

    const mockSheet = {
      name: "TestSheet",
      getRangeByIndexes: jest.fn(),
      delete: jest.fn(),
    };

    const mockRange = {
      values: [["METADATA", "1234567890000"]],
      load: jest.fn(),
    };

    beforeEach(() => {
      jest.clearAllMocks();
      
      // Mock Excel.run
      (global as any).Excel = {
        run: jest.fn((callback) => callback(mockContext)),
      };

      mockContext.workbook.worksheets.items = [mockSheet];
      mockContext.workbook.worksheets.getItem.mockReturnValue(mockSheet);
      mockSheet.getRangeByIndexes.mockReturnValue(mockRange);
    });

    describe("getSheetMetadata", () => {
      it("should return metadata when it exists", async () => {
        mockRange.values = [["METADATA", "1234567890000"]];
        
        const metadata = await getSheetMetadata("TestSheet");
        
        expect(metadata).toEqual({
          timestamp: 1234567890000,
        });
      });

      it("should return null when metadata marker is missing", async () => {
        mockRange.values = [["NOT_METADATA", "1234567890000"]];
        
        const metadata = await getSheetMetadata("TestSheet");
        
        expect(metadata).toBeNull();
      });

      it("should return null on error", async () => {
        (global as any).Excel.run = jest.fn(() => Promise.reject(new Error("Sheet not found")));
        
        const metadata = await getSheetMetadata("NonExistentSheet");
        
        expect(metadata).toBeNull();
      });
    });

    describe("setSheetMetadata", () => {
      it("should write metadata to row 0", async () => {
        const metadata: SheetMetadata = {
          timestamp: 1234567890000,
        };

        await setSheetMetadata("TestSheet", metadata);

        expect(mockSheet.getRangeByIndexes).toHaveBeenCalledWith(0, 0, 1, 2);
        expect(mockRange.values).toEqual([["METADATA", "1234567890000"]]);
      });
    });

    describe("isSheetDataStale", () => {
      it("should return true when metadata is missing", async () => {
        mockRange.values = [["NOT_METADATA", "1234567890000"]];
        
        const isStale = await isSheetDataStale("TestSheet");
        
        expect(isStale).toBe(true);
      });

      it("should return true when data is older than 2 days", async () => {
        const threeDaysAgo = Date.now() - (3 * 24 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", threeDaysAgo.toString()]];
        
        const isStale = await isSheetDataStale("TestSheet");
        
        expect(isStale).toBe(true);
      });

      it("should return false when data is less than 2 days old", async () => {
        const oneDayAgo = Date.now() - (1 * 24 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", oneDayAgo.toString()]];
        
        const isStale = await isSheetDataStale("TestSheet");
        
        expect(isStale).toBe(false);
      });

      it("should return true on error", async () => {
        (global as any).Excel.run = jest.fn(() => Promise.reject(new Error("Error")));
        
        const isStale = await isSheetDataStale("TestSheet");
        
        expect(isStale).toBe(true);
      });
    });

    describe("sheetExists", () => {
      it("should return true when sheet exists", async () => {
        mockContext.workbook.worksheets.items = [{ name: "TestSheet" }];
        
        const exists = await sheetExists("TestSheet");
        
        expect(exists).toBe(true);
      });

      it("should return false when sheet does not exist", async () => {
        mockContext.workbook.worksheets.items = [{ name: "OtherSheet" }];
        
        const exists = await sheetExists("TestSheet");
        
        expect(exists).toBe(false);
      });

      it("should return false on error", async () => {
        (global as any).Excel.run = jest.fn(() => Promise.reject(new Error("Error")));
        
        const exists = await sheetExists("TestSheet");
        
        expect(exists).toBe(false);
      });
    });

    describe("deleteSheetIfExists", () => {
      it("should delete sheet when it exists", async () => {
        mockContext.workbook.worksheets.items = [mockSheet];
        
        await deleteSheetIfExists("TestSheet");
        
        expect(mockSheet.delete).toHaveBeenCalled();
      });

      it("should not throw error when sheet does not exist", async () => {
        mockContext.workbook.worksheets.items = [];
        
        await expect(deleteSheetIfExists("TestSheet")).resolves.not.toThrow();
      });
    });

    describe("refreshSheetIfStale", () => {
      const mockRecreateFunction = jest.fn().mockResolvedValue(undefined);

      beforeEach(() => {
        mockRecreateFunction.mockClear();
      });

      it("should return false when sheet does not exist", async () => {
        mockContext.workbook.worksheets.items = [];
        
        const refreshed = await refreshSheetIfStale("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(false);
        expect(mockRecreateFunction).not.toHaveBeenCalled();
      });

      it("should refresh when sheet exists and is stale", async () => {
        mockContext.workbook.worksheets.items = [mockSheet];
        const threeDaysAgo = Date.now() - (3 * 24 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", threeDaysAgo.toString()]];
        
        const refreshed = await refreshSheetIfStale("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(true);
        expect(mockSheet.delete).toHaveBeenCalled();
        expect(mockRecreateFunction).toHaveBeenCalled();
      });

      it("should not refresh when sheet exists but is fresh", async () => {
        mockContext.workbook.worksheets.items = [mockSheet];
        const oneDayAgo = Date.now() - (1 * 24 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", oneDayAgo.toString()]];
        
        const refreshed = await refreshSheetIfStale("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(false);
        expect(mockSheet.delete).not.toHaveBeenCalled();
        expect(mockRecreateFunction).not.toHaveBeenCalled();
      });
    });

    describe("refreshSheetOnLogin", () => {
      const mockRecreateFunction = jest.fn().mockResolvedValue(undefined);

      beforeEach(() => {
        mockRecreateFunction.mockClear();
      });

      it("should return false when sheet does not exist", async () => {
        mockContext.workbook.worksheets.items = [];
        
        const refreshed = await refreshSheetOnLogin("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(false);
        expect(mockRecreateFunction).not.toHaveBeenCalled();
      });

      it("should always refresh when sheet exists (regardless of age)", async () => {
        mockContext.workbook.worksheets.items = [mockSheet];
        const oneDayAgo = Date.now() - (1 * 24 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", oneDayAgo.toString()]];
        
        const refreshed = await refreshSheetOnLogin("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(true);
        expect(mockSheet.delete).toHaveBeenCalled();
        expect(mockRecreateFunction).toHaveBeenCalled();
      });

      it("should refresh even with fresh data (< 2 days)", async () => {
        mockContext.workbook.worksheets.items = [mockSheet];
        const oneHourAgo = Date.now() - (1 * 60 * 60 * 1000);
        mockRange.values = [["METADATA", oneHourAgo.toString()]];
        
        const refreshed = await refreshSheetOnLogin("TestSheet", mockRecreateFunction);
        
        expect(refreshed).toBe(true);
        expect(mockSheet.delete).toHaveBeenCalled();
        expect(mockRecreateFunction).toHaveBeenCalled();
      });
    });
  });
});
