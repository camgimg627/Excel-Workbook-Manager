import { createSeedDemoSession } from "@/lib/demo-data";
import { acceptCapture, addInventory, cookRecipe, purchaseShoppingItem } from "@/lib/session-actions";

describe("session actions", () => {
  it("adds manual inventory items into the household state", () => {
    const state = createSeedDemoSession();
    const nextState = addInventory(state, {
      name: "Dishwasher Pods",
      category: "household_supplies",
      subcategory: "cleaning",
      quantityValue: 24,
      quantityUnit: "count",
      quantityFullValue: 24,
      spaceId: "space-rivera-garage",
      visibility: "shared",
      notes: "Monthly warehouse refill",
      retailer: "Manual log",
    });

    expect(nextState.inventoryItems.some((entry) => entry.productId !== undefined && entry.notes === "Monthly warehouse refill")).toBe(true);
    expect(nextState.purchaseEvents.length).toBe(state.purchaseEvents.length + 1);
  });

  it("accepts a receipt draft and merges it into inventory plus purchase history", () => {
    const state = createSeedDemoSession();
    const draftId = state.captureDrafts[0].id;
    const nextState = acceptCapture(state, draftId);
    const paperTowels = nextState.inventoryItems.find((entry) => entry.id === "item-rivera-paper-towels");

    expect(nextState.captureDrafts).toHaveLength(0);
    expect(paperTowels?.quantityValue).toBe(7);
    expect(nextState.purchaseEvents.length).toBe(state.purchaseEvents.length + 2);
  });

  it("logs a recipe cook and deducts matching ingredients", () => {
    const state = createSeedDemoSession();
    const nextState = cookRecipe(state, "recipe-rivera-spaghetti", 4);
    const spaghetti = nextState.inventoryItems.find((entry) => entry.id === "item-rivera-spaghetti");
    const sauce = nextState.inventoryItems.find((entry) => entry.id === "item-rivera-pasta-sauce");

    expect(nextState.recipeLogs.length).toBe(state.recipeLogs.length + 1);
    expect(spaghetti?.quantityValue).toBe(0);
    expect(sauce?.quantityValue).toBe(0);
    expect(nextState.consumptionEvents.length).toBeGreaterThan(state.consumptionEvents.length);
  });

  it("marks a shopping-list item purchased and restocks inventory", () => {
    const state = createSeedDemoSession();
    const entry = state.shoppingListEntries.find((item) => item.productName === "Fruit Snacks");
    expect(entry).toBeDefined();

    const nextState = purchaseShoppingItem(state, entry!);
    const newInventory = nextState.inventoryItems.find((item) => {
      const product = nextState.products.find((productEntry) => productEntry.id === item.productId);
      return product?.name === "Fruit Snacks";
    });

    expect(nextState.shoppingListEntries.some((item) => item.id === entry!.id)).toBe(false);
    expect(newInventory).toBeDefined();
    expect(nextState.purchaseEvents.length).toBe(state.purchaseEvents.length + 1);
  });
});
