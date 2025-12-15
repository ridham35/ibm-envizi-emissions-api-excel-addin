// Copyright IBM Corp. 2025
import * as api from "../src/functions/functions"
import { genericApiCall } from "../src/functions/generic-api-call";
import { factorSearch } from "../src/functions/factorSearchHelper";
import { factorHelper } from "../src/functions/factorHelper";
import { handleTypesFunction } from "../src/functions/types-handler";

jest.mock("../src/functions/generic-api-call", () => ({
  genericApiCall: jest.fn(),
}));

jest.mock("../src/functions/factorSearchHelper", () => ({
  factorSearch: jest.fn(),
}));

jest.mock("../src/functions/factorHelper", () => ({
  factorHelper: jest.fn(),
}));

jest.mock("../src/functions/types-handler", () => ({
  handleTypesFunction: jest.fn(),
}));


describe("emissions custom functions", () => {
  const mockedGenericApiCall = genericApiCall as jest.MockedFunction<typeof genericApiCall>;

  beforeEach(() => {
    mockedGenericApiCall.mockResolvedValue([["mocked-response"]]);
  });

  const commonTests: {
    fn: (...args: any[]) => Promise<any[][]>;
    name: string;
    apiType: Parameters<typeof genericApiCall>[0];
    args: any[];
    expectedPayload: any;
  }[] = [
    {
      fn: api.location,
      name: "location",
      apiType: "location",
      args: ["electricity", 100, "kWh","USA", "NY", "2024-01-01", "grid1"],
      expectedPayload: {
        type: "electricity",
        value: 100,
        country: "USA",
        stateProvince: "NY",
        date: "2024-01-01",
        unit: "kWh",
        powerGrid: "grid1",
      },
    },
    {
      fn: api.location_by_factorId,
      name: "location_by_factorId",
      apiType: "location",
      args: [123, 200, "kg"],
      expectedPayload: { factorId: 123, value: 200, unit: "kg" },
    },
    {
      fn: api.stationary,
      name: "stationary",
      apiType: "stationary",
      args: ["fuel", 10, "L", "USA", "CA", "2024-01-01"],
      expectedPayload: {
        type: "fuel",
        value: 10,
        unit: "L",
        country: "USA",
        stateProvince: "CA",
        date: "2024-01-01",
      },
    },
    {
      fn: api.stationary_by_factorId,
      name: "stationary_by_factorId",
      apiType: "stationary",
      args: [456, 20, "L"],
      expectedPayload: { factorId: 456, value: 20, unit: "L" },
    },
    {
      fn: api.fugitive,
      name: "fugitive",
      apiType: "fugitive",
      args: ["gas", 50, "kg", "USA", "TX", "2024-01-01"],
      expectedPayload: {
        type: "gas",
        value: 50,
        unit: "kg",
        country: "USA",
        stateProvince: "TX",
        date: "2024-01-01",
      },
    },
    {
      fn: api.fugitive_by_factorId,
      name: "fugitive_by_factorId",
      apiType: "fugitive",
      args: [789, 30, "kg"],
      expectedPayload: { factorId: 789, value: 30, unit: "kg" },
    },
    {
      fn: api.mobile,
      name: "mobile",
      apiType: "mobile",
      args: ["diesel", 70, "L", "USA", "FL", "2024-01-01"],
      expectedPayload: {
        type: "diesel",
        value: 70,
        unit: "L",
        country: "USA",
        stateProvince: "FL",
        date: "2024-01-01",
      },
    },
    {
      fn: api.mobile_by_factorId,
      name: "mobile_by_factorId",
      apiType: "mobile",
      args: [111, 90, "L"],
      expectedPayload: { factorId: 111, value: 90, unit: "L" },
    },
    {
      fn: api.transportation_and_distribution,
      name: "transportation_and_distribution",
      apiType: "transportation_and_distribution",
      args: ["shipping", 5, "ton", "USA", "WA", "2024-01-01"],
      expectedPayload: {
        type: "shipping",
        value: 5,
        unit: "ton",
        country: "USA",
        stateProvince: "WA",
        date: "2024-01-01",
      },
    },
    {
      fn: api.transportation_and_distribution_by_factorId,
      name: "transportation_and_distribution_by_factorId",
      apiType: "transportation_and_distribution",
      args: [222, 15, "ton"],
      expectedPayload: { factorId: 222, value: 15, unit: "ton" },
    },
    {
      fn: api.calculation,
      name: "calculation",
      apiType: "calculation",
      args: ["custom", 999, "kg", "USA", "NJ", "2024-01-01", "gridX"],
      expectedPayload: {
        type: "custom",
        value: 999,
        unit: "kg",
        country: "USA",
        stateProvince: "NJ",
        date: "2024-01-01",
        powerGrid: "gridX",
      },
    },
    {
      fn: api.calculation_by_factorId,
      name: "calculation_by_factorId",
      apiType: "calculation",
      args: [333, 444, "kg"],
      expectedPayload: { factorId: 333, value: 444, unit: "kg" },
    },
  ];

  test.each(commonTests)("$name calls genericApiCall correctly", async ({ fn, apiType, args, expectedPayload }) => {
    const result = await fn(...args);

    expect(mockedGenericApiCall).toHaveBeenCalledWith(apiType, expectedPayload);
    expect(result).toEqual([["mocked-response"]]);
  });

  it("propagates errors from genericApiCall", async () => {
    mockedGenericApiCall.mockRejectedValueOnce(new Error("Boom!"));
    await expect(api.location("electricity", 1, "kwh" , "USA")).rejects.toThrow("Boom!");
  });
});

describe("factor-related functions", () => {
  const mockedFactorSearch = factorSearch as jest.MockedFunction<typeof factorSearch>;
  const mockedFactorHelper = factorHelper as jest.MockedFunction<typeof factorHelper>;

  beforeEach(() => {
    mockedFactorSearch.mockResolvedValue([["search-result"]]);
    mockedFactorHelper.mockResolvedValue([["helper-result"]]);
  });

  it("factor_search calls factorSearch with correct arguments", async () => {
    const result = await api.factor_search("electricity", "usa","new york","10/10/2020");
    expect(mockedFactorSearch).toHaveBeenCalledWith("electricity", "usa","new york","10/10/2020", undefined, undefined);
    expect(result).toEqual([["search-result"]]);
  });

  it("factor_search calls factorSearch with pagination parameters", async () => {
    const result = await api.factor_search("electricity", "usa", "new york", "10/10/2020", 2, 50);
    expect(mockedFactorSearch).toHaveBeenCalledWith("electricity", "usa", "new york", "10/10/2020", 2, 50);
    expect(result).toEqual([["search-result"]]);
  });

  it("factor calls factorHelper with correct arguments", async () => {
    const result = await api.factor("electricity", "kwh", "usa","new york","10/10/2020");
    expect(mockedFactorHelper).toHaveBeenCalledWith("electricity", "kwh", "usa","new york","10/10/2020");
    expect(result).toEqual([["helper-result"]]);
  });

  it("factorById calls factorHelper with correct arguments", async () => {
    const result = await api.factor_by_id(123, "kg");
    expect(mockedFactorHelper).toHaveBeenCalledWith(123, "kg");
    expect(result).toEqual([["helper-result"]]);
  });
});

describe("types function", () => {
  const mockedHandleTypesFunction = handleTypesFunction as jest.MockedFunction<typeof handleTypesFunction>;

  beforeEach(() => {
    mockedHandleTypesFunction.mockResolvedValue("location");
  });

  it("should call handleTypesFunction with correct arguments", async () => {
    const mockInvocation = {
      address: "Sheet1!A1",
    } as CustomFunctions.Invocation;

    const result = await api.types("location", mockInvocation);

    expect(mockedHandleTypesFunction).toHaveBeenCalledWith("location", mockInvocation);
    expect(result).toBe("location");
  });

  it("should handle different API names", async () => {
    const mockInvocation = {
      address: "Sheet1!B2",
    } as CustomFunctions.Invocation;

    mockedHandleTypesFunction.mockResolvedValue("mobile");
    const result = await api.types("mobile", mockInvocation);

    expect(mockedHandleTypesFunction).toHaveBeenCalledWith("mobile", mockInvocation);
    expect(result).toBe("mobile");
  });

  it("should handle factor API name", async () => {
    const mockInvocation = {
      address: "Sheet1!C3",
    } as CustomFunctions.Invocation;

    mockedHandleTypesFunction.mockResolvedValue("factor");
    const result = await api.types("factor", mockInvocation);

    expect(mockedHandleTypesFunction).toHaveBeenCalledWith("factor", mockInvocation);
    expect(result).toBe("factor");
  });

  it("should propagate errors from handleTypesFunction", async () => {
    const mockInvocation = {
      address: "Sheet1!A1",
    } as CustomFunctions.Invocation;

    mockedHandleTypesFunction.mockRejectedValue(new Error("Validation error"));

    await expect(api.types("invalid", mockInvocation)).rejects.toThrow("Validation error");
  });

  it("should pass invocation object correctly", async () => {
    const mockInvocation = {
      address: "Sheet2!D10",
      parameterAddresses: [],
    } as CustomFunctions.Invocation;

    await api.types("calculation", mockInvocation);

    expect(mockedHandleTypesFunction).toHaveBeenCalledWith("calculation", mockInvocation);
  });
});
