import { addDays } from "@/lib/date";
import { Category, DemoSessionState, ForecastSnapshot, InventoryItem, QuantityUnit } from "@/lib/domain";
import { round } from "@/lib/utils";

const defaultRates: Record<Category, Record<QuantityUnit, number>> = {
  food: {
    count: 0.35,
    ml: 90,
    g: 55,
    percentage: 10,
  },
  household_supplies: {
    count: 0.2,
    ml: 45,
    g: 25,
    percentage: 4.5,
  },
  personal_care: {
    count: 0.12,
    ml: 20,
    g: 12,
    percentage: 5.5,
  },
};

function scheduleUsesPerDay(scheduleFrequency: "daily" | "weekly" | "custom", frequencyDays: number[]): number {
  if (scheduleFrequency === "daily") {
    return 1;
  }
  if (scheduleFrequency === "weekly") {
    return 1 / 7;
  }
  return Math.max(frequencyDays.length, 1) / 7;
}

function getHistoricalRate(state: DemoSessionState, item: InventoryItem): number | null {
  const events = state.consumptionEvents
    .filter((entry) => entry.inventoryItemId === item.id)
    .sort((left, right) => new Date(left.consumedAt).getTime() - new Date(right.consumedAt).getTime());

  if (!events.length) {
    return null;
  }

  const totalUsed = events.reduce((sum, entry) => sum + entry.quantityUsed, 0);
  const firstDate = new Date(events[0].consumedAt);
  const lastDate = new Date(events[events.length - 1].consumedAt);
  const daysSpan = Math.max((lastDate.getTime() - firstDate.getTime()) / (1000 * 60 * 60 * 24), 1);

  return totalUsed / daysSpan;
}

function getDailyConsumptionRate(state: DemoSessionState, item: InventoryItem): { rate: number; confidence: ForecastSnapshot["confidence"] } {
  const product = state.products.find((entry) => entry.id === item.productId);
  const category = product?.category ?? "food";
  const matchingSchedule = state.consumptionSchedules.find(
    (entry) => entry.householdId === item.householdId && entry.productId === item.productId
  );

  if (matchingSchedule) {
    return {
      rate: matchingSchedule.quantityPerUse * scheduleUsesPerDay(matchingSchedule.frequency, matchingSchedule.frequencyDays),
      confidence: "scheduled",
    };
  }

  const historicalRate = getHistoricalRate(state, item);
  if (historicalRate && historicalRate > 0) {
    return {
      rate: historicalRate,
      confidence: "historical",
    };
  }

  return {
    rate: defaultRates[category][item.quantityUnit],
    confidence: "default",
  };
}

export function getStockPercent(item: InventoryItem): number {
  if (item.quantityFullValue <= 0) {
    return 0;
  }

  return Math.min(item.quantityValue / item.quantityFullValue, 1);
}

export function getForecastSnapshots(state: DemoSessionState, householdId: string, inventory: InventoryItem[]): ForecastSnapshot[] {
  const household = state.households.find((entry) => entry.id === householdId);
  const leadDays = household?.settings.replenishmentLeadDays ?? 7;
  const alertLeadDays = household?.settings.alertLeadDays ?? 4;

  return inventory
    .map((item) => {
      const product = state.products.find((entry) => entry.id === item.productId);
      const { rate, confidence } = getDailyConsumptionRate(state, item);
      const safeRate = Math.max(rate, 0.01);
      const daysRemaining = round(item.quantityValue / safeRate, 1);
      const replenishmentDate = addDays(new Date(), Math.max(Math.ceil(daysRemaining), 0)).toISOString();
      const alertDate = addDays(replenishmentDate, -alertLeadDays).toISOString();
      const stockPercent = getStockPercent(item);

      return {
        inventoryItemId: item.id,
        householdId: item.householdId,
        productId: item.productId,
        productName: product?.name ?? "Unknown product",
        category: product?.category ?? "food",
        quantityUnit: item.quantityUnit,
        quantityValue: item.quantityValue,
        quantityFullValue: item.quantityFullValue,
        dailyConsumptionRate: round(safeRate, 2),
        daysRemaining,
        replenishmentDate,
        alertDate,
        lowStock: stockPercent <= 0.25 || daysRemaining <= leadDays,
        confidence,
      } satisfies ForecastSnapshot;
    })
    .sort((left, right) => left.daysRemaining - right.daysRemaining);
}
