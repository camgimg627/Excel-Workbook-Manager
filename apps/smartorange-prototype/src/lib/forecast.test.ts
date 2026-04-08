import { createSeedDemoSession } from "@/lib/demo-data";
import { getForecastSnapshots } from "@/lib/forecast";
import { getVisibleInventoryItems } from "@/lib/selectors";

describe("forecast logic", () => {
  it("flags low-stock count items using scheduled depletion", () => {
    const state = createSeedDemoSession();
    const inventory = getVisibleInventoryItems(state, "household-rivera", "member-rivera-owner");
    const snapshots = getForecastSnapshots(state, "household-rivera", inventory);
    const paperTowels = snapshots.find((entry) => entry.productName === "Paper Towels");

    expect(paperTowels).toBeDefined();
    expect(paperTowels?.lowStock).toBe(true);
    expect(paperTowels?.daysRemaining).toBeGreaterThanOrEqual(6.5);
    expect(paperTowels?.daysRemaining).toBeLessThanOrEqual(7.5);
  });

  it("supports percentage-based depletion forecasting", () => {
    const state = createSeedDemoSession();
    const inventory = getVisibleInventoryItems(state, "household-maple", "member-maple-owner");
    const snapshots = getForecastSnapshots(state, "household-maple", inventory);
    const dishSoap = snapshots.find((entry) => entry.productName === "Dish Soap");

    expect(dishSoap).toBeDefined();
    expect(dishSoap?.quantityUnit).toBe("percentage");
    expect(dishSoap?.lowStock).toBe(true);
    expect(dishSoap?.daysRemaining).toBeCloseTo(9, 0);
  });
});
