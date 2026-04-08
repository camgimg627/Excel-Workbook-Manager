import { daysUntil } from "@/lib/date";
import {
  DemoSessionState,
  ForecastSnapshot,
  Household,
  HouseholdMember,
  InventoryItem,
  MemberRole,
  NotificationItem,
  ShoppingListEntry,
  Space,
} from "@/lib/domain";
import { getForecastSnapshots } from "@/lib/forecast";
import { createId, round } from "@/lib/utils";

export function getActiveHousehold(state: DemoSessionState): Household {
  return state.households.find((entry) => entry.id === state.selectedHouseholdId) ?? state.households[0];
}

export function getHouseholdMembers(state: DemoSessionState, householdId: string): HouseholdMember[] {
  return state.householdMembers.filter((entry) => entry.householdId === householdId);
}

export function getActiveMember(state: DemoSessionState): HouseholdMember {
  const household = getActiveHousehold(state);

  return (
    state.householdMembers.find(
      (entry) => entry.id === state.selectedMemberId && entry.householdId === household.id
    ) ??
    getHouseholdMembers(state, household.id)[0]
  );
}

export function getActiveUser(state: DemoSessionState) {
  const activeMember = getActiveMember(state);
  return state.users.find((entry) => entry.id === activeMember.userId) ?? state.users[0];
}

export function getSpacesForHousehold(state: DemoSessionState, householdId: string): Space[] {
  return state.spaces.filter((entry) => entry.householdId === householdId);
}

export function canViewItem(item: InventoryItem, role: MemberRole, activeUserId: string): boolean {
  if (role === "owner") {
    return true;
  }
  if (role === "member") {
    return item.visibility !== "private" || item.ownerUserId === activeUserId;
  }

  return item.visibility !== "private";
}

export function getVisibleInventoryItems(state: DemoSessionState, householdId: string, memberId: string): InventoryItem[] {
  const activeMember = state.householdMembers.find((entry) => entry.id === memberId);
  if (!activeMember) {
    return [];
  }

  return state.inventoryItems
    .filter((entry) => entry.householdId === householdId)
    .filter((entry) => canViewItem(entry, activeMember.role, activeMember.userId));
}

export function getForecastForActiveView(state: DemoSessionState): ForecastSnapshot[] {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const inventory = getVisibleInventoryItems(state, household.id, member.id);
  return getForecastSnapshots(state, household.id, inventory);
}

export function getExpiringItems(state: DemoSessionState): InventoryItem[] {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  return getVisibleInventoryItems(state, household.id, member.id)
    .filter((entry) => entry.expiryDate && daysUntil(entry.expiryDate) <= 5)
    .sort((left, right) => daysUntil(left.expiryDate ?? "") - daysUntil(right.expiryDate ?? ""));
}

export function getCurrentShoppingList(state: DemoSessionState): ShoppingListEntry[] {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const visibleInventory = getVisibleInventoryItems(state, household.id, member.id);
  const activeManualEntries = state.shoppingListEntries.filter((entry) => entry.householdId === household.id);
  const existingKeys = new Set(
    activeManualEntries.map((entry) => `${entry.inventoryItemId ?? "manual"}:${entry.productName.toLowerCase()}`)
  );
  const autoEntries: ShoppingListEntry[] = [];
  const forecasts = getForecastSnapshots(state, household.id, visibleInventory);

  forecasts
    .filter((entry) => entry.lowStock)
    .forEach((entry) => {
      const item = visibleInventory.find((inventoryItem) => inventoryItem.id === entry.inventoryItemId);
      if (!item) {
        return;
      }

      const key = `${entry.inventoryItemId}:${entry.productName.toLowerCase()}`;
      if (existingKeys.has(key)) {
        return;
      }

      autoEntries.push({
        id: createId("auto-shopping"),
        householdId: household.id,
        inventoryItemId: entry.inventoryItemId,
        productId: entry.productId,
        productName: entry.productName,
        category: entry.category,
        quantity: Math.max(round(item.quantityFullValue - item.quantityValue, 0), 1),
        quantityUnit: item.quantityUnit,
        spaceId: item.spaceId,
        source: "forecast",
        note: `Projected to run out in ${entry.daysRemaining} days.`,
        createdBy: member.userId,
        createdAt: new Date().toISOString(),
      });
    });

  getExpiringItems(state).forEach((item) => {
    const product = state.products.find((entry) => entry.id === item.productId);
    const key = `${item.id}:${product?.name.toLowerCase() ?? "unknown"}`;
    if (existingKeys.has(key)) {
      return;
    }

    autoEntries.push({
      id: createId("expiry-shopping"),
      householdId: household.id,
      inventoryItemId: item.id,
      productId: item.productId,
      productName: product?.name ?? "Unknown product",
      category: product?.category ?? "food",
      quantity: Math.max(round(item.quantityFullValue - item.quantityValue, 0), 1),
      quantityUnit: item.quantityUnit,
      spaceId: item.spaceId,
      source: "expiry",
      note: `Expires ${item.expiryDate}. Replace before it falls out of rotation.`,
      createdBy: member.userId,
      createdAt: new Date().toISOString(),
    });
  });

  return [...autoEntries, ...activeManualEntries];
}

export function getNotifications(state: DemoSessionState): NotificationItem[] {
  const household = getActiveHousehold(state);
  const forecasts = getForecastForActiveView(state).slice(0, 3);
  const expiring = getExpiringItems(state).slice(0, 2);

  const notifications: NotificationItem[] = forecasts.map((entry) => ({
    id: `notification-${entry.inventoryItemId}`,
    householdId: household.id,
    severity: entry.daysRemaining <= 3 ? "urgent" : "warning",
    title: `${entry.productName} is trending low`,
    body: `At the current pace, it will need attention in about ${entry.daysRemaining} days.`,
    routeHint: "/shopping-list",
  }));

  expiring.forEach((item) => {
    const product = state.products.find((entry) => entry.id === item.productId);
    notifications.push({
      id: `notification-expiry-${item.id}`,
      householdId: household.id,
      severity: "warning",
      title: `${product?.name ?? "Item"} expires soon`,
      body: `${product?.name ?? "Item"} should be used or replaced within ${Math.max(daysUntil(item.expiryDate ?? ""), 0)} days.`,
      routeHint: "/inventory",
    });
  });

  return notifications.slice(0, 5);
}

export function getRecentActivity(state: DemoSessionState): Array<{ id: string; label: string; detail: string; at: string }> {
  const household = getActiveHousehold(state);
  const activity = [
    ...state.purchaseEvents
      .filter((entry) => entry.householdId === household.id)
      .map((entry) => {
        const product = state.products.find((productEntry) => productEntry.id === entry.productId);
        return {
          id: entry.id,
          label: `Purchased ${product?.name ?? "item"}`,
          detail: `${entry.retailer} • ${entry.quantityPurchased} ${entry.quantityUnit}`,
          at: entry.purchasedAt,
        };
      }),
    ...state.consumptionEvents
      .map((entry) => {
        const item = state.inventoryItems.find((inventoryEntry) => inventoryEntry.id === entry.inventoryItemId);
        if (!item || item.householdId !== household.id) {
          return null;
        }
        const product = state.products.find((productEntry) => productEntry.id === item.productId);
        return {
          id: entry.id,
          label: `Used ${product?.name ?? "item"}`,
          detail: `${entry.quantityUsed} ${item.quantityUnit} logged ${entry.source}`,
          at: entry.consumedAt,
        };
      })
      .filter((entry): entry is { id: string; label: string; detail: string; at: string } => Boolean(entry)),
    ...state.recipeLogs
      .filter((entry) => entry.householdId === household.id)
      .map((entry) => {
        const recipe = state.recipes.find((recipeEntry) => recipeEntry.id === entry.recipeId);
        return {
          id: entry.id,
          label: `Cooked ${recipe?.name ?? "recipe"}`,
          detail: `${entry.servingsMade} servings`,
          at: entry.cookedAt,
        };
      }),
  ];

  return activity
    .sort((left, right) => new Date(right.at).getTime() - new Date(left.at).getTime())
    .slice(0, 8);
}

export function getSpaceSnapshots(state: DemoSessionState): Array<{ space: Space; itemCount: number; lowStockCount: number }> {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const visibleInventory = getVisibleInventoryItems(state, household.id, member.id);
  const forecasts = getForecastSnapshots(state, household.id, visibleInventory);

  return getSpacesForHousehold(state, household.id).map((space) => {
    const inSpace = visibleInventory.filter((entry) => entry.spaceId === space.id);
    const lowStockCount = forecasts.filter((entry) => inSpace.some((item) => item.id === entry.inventoryItemId) && entry.lowStock).length;

    return {
      space,
      itemCount: inSpace.length,
      lowStockCount,
    };
  });
}

export function getRolePermissions(role: MemberRole) {
  return {
    canEditInventory: role !== "contributor",
    canCaptureToInventory: role !== "contributor",
    canPurchaseShopping: role !== "contributor",
    canCookRecipes: role !== "contributor",
    canAddShoppingItems: true,
    canResetDemo: role === "owner",
  };
}

export function getDashboardMetrics(state: DemoSessionState) {
  const visibleInventory = getVisibleInventoryItems(state, getActiveHousehold(state).id, getActiveMember(state).id);
  const forecasts = getForecastForActiveView(state);
  const lowStockCount = forecasts.filter((entry) => entry.lowStock).length;
  const expiringCount = getExpiringItems(state).length;
  const sharedCount = visibleInventory.filter((entry) => entry.visibility !== "private").length;
  const averageDays = forecasts.length
    ? round(forecasts.reduce((sum, entry) => sum + entry.daysRemaining, 0) / forecasts.length, 1)
    : 0;

  return {
    visibleItemCount: visibleInventory.length,
    lowStockCount,
    expiringCount,
    sharedCount,
    averageDays,
  };
}
