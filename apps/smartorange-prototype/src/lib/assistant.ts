import { formatDate } from "@/lib/date";
import { DemoSessionState } from "@/lib/domain";
import {
  getActiveHousehold,
  getActiveMember,
  getCurrentShoppingList,
  getExpiringItems,
  getForecastForActiveView,
  getVisibleInventoryItems,
} from "@/lib/selectors";

function findRelevantItem(state: DemoSessionState, prompt: string) {
  const lowerPrompt = prompt.toLowerCase();
  const visibleInventory = getVisibleInventoryItems(state, getActiveHousehold(state).id, getActiveMember(state).id);

  return visibleInventory.find((item) => {
    const product = state.products.find((entry) => entry.id === item.productId);
    return product ? lowerPrompt.includes(product.name.toLowerCase()) : false;
  });
}

export function buildAssistantReply(state: DemoSessionState, prompt: string): string {
  const lowerPrompt = prompt.toLowerCase();
  const household = getActiveHousehold(state);
  const forecasts = getForecastForActiveView(state);
  const shoppingList = getCurrentShoppingList(state);
  const expiring = getExpiringItems(state);
  const relevantItem = findRelevantItem(state, prompt);

  if (lowerPrompt.includes("run low") || lowerPrompt.includes("running low") || lowerPrompt.includes("shopping list")) {
    const topEntries = shoppingList.slice(0, 3);
    if (!topEntries.length) {
      return `${household.name} does not have any active replenishment calls right now. Shared staples are covering the next few days comfortably.`;
    }

    return topEntries
      .map((entry, index) => `${index + 1}. ${entry.productName}: ${entry.note ?? "needs restocking soon"}`)
      .join(" ");
  }

  if (lowerPrompt.includes("expire")) {
    if (!expiring.length) {
      return `${household.name} has no visible items expiring within the next five days.`;
    }

    return expiring
      .slice(0, 3)
      .map((item) => {
        const product = state.products.find((entry) => entry.id === item.productId);
        return `${product?.name ?? "Item"} expires ${formatDate(item.expiryDate ?? "")}`;
      })
      .join(". ");
  }

  if (lowerPrompt.includes("last buy") || lowerPrompt.includes("when did we last buy")) {
    const item = relevantItem;
    if (!item) {
      return "Name the product and I can look up the latest purchase event in this household.";
    }

    const purchase = [...state.purchaseEvents]
      .filter((entry) => entry.householdId === household.id && entry.productId === item.productId)
      .sort((left, right) => new Date(right.purchasedAt).getTime() - new Date(left.purchasedAt).getTime())[0];

    if (!purchase) {
      return "I do not have a recorded purchase for that item in this demo household yet.";
    }

    const product = state.products.find((entry) => entry.id === item.productId);
    return `${product?.name ?? "That item"} was last purchased on ${formatDate(purchase.purchasedAt)} from ${purchase.retailer}.`;
  }

  if (lowerPrompt.includes("enough") || lowerPrompt.includes("do we have")) {
    const item = relevantItem;
    if (!item) {
      return "Ask about a specific product and I will compare current quantity with forecasted usage.";
    }

    const product = state.products.find((entry) => entry.id === item.productId);
    const forecast = forecasts.find((entry) => entry.inventoryItemId === item.id);
    if (!forecast) {
      return `${product?.name ?? "That item"} is visible, but I do not have enough signal to forecast it yet.`;
    }

    if (forecast.daysRemaining >= 3) {
      return `Yes. ${product?.name ?? "That item"} has about ${forecast.daysRemaining} days of coverage left at the current pace.`;
    }

    return `Barely. ${product?.name ?? "That item"} only has about ${forecast.daysRemaining} days of coverage left, so it should move onto the shopping list.`;
  }

  if (lowerPrompt.includes("where")) {
    const item = relevantItem;
    if (!item) {
      return "Ask where a specific product lives and I will map it to the household space in this prototype.";
    }

    const product = state.products.find((entry) => entry.id === item.productId);
    const space = state.spaces.find((entry) => entry.id === item.spaceId);
    return `${product?.name ?? "That item"} is currently tracked in ${space?.name ?? "an unknown space"}.`;
  }

  const topForecast = forecasts[0];
  if (!topForecast) {
    return `${household.name} has no visible inventory yet, so the next best step is to log or capture a few core items.`;
  }

  return `${household.name} currently shows ${forecasts.filter((entry) => entry.lowStock).length} low-stock items. ${topForecast.productName} is the next likely replenishment need at roughly ${topForecast.daysRemaining} days remaining.`;
}
