import { createSeedDemoSession } from "@/lib/demo-data";
import { getCurrentShoppingList, getVisibleInventoryItems } from "@/lib/selectors";
import { switchMember } from "@/lib/session-actions";

describe("visibility and shopping selectors", () => {
  it("hides private items from contributors while owners can still see them", () => {
    const ownerState = createSeedDemoSession();
    const ownerInventory = getVisibleInventoryItems(ownerState, "household-rivera", "member-rivera-owner");
    expect(ownerInventory.some((entry) => entry.id === "item-rivera-allergy")).toBe(true);

    const contributorState = switchMember(ownerState, "member-rivera-contributor");
    const contributorInventory = getVisibleInventoryItems(
      contributorState,
      "household-rivera",
      "member-rivera-contributor"
    );
    expect(contributorInventory.some((entry) => entry.id === "item-rivera-allergy")).toBe(false);
  });

  it("combines manual and auto-generated shopping list entries", () => {
    const state = createSeedDemoSession();
    const entries = getCurrentShoppingList(state);

    expect(entries.some((entry) => entry.productName === "Fruit Snacks" && entry.source === "manual")).toBe(true);
    expect(entries.some((entry) => entry.productName === "Milk" && entry.source === "forecast")).toBe(true);
  });
});
