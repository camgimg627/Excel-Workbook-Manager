import { Category, QuantityUnit, Visibility } from "@/lib/domain";

export const categoryLabels: Record<Category, string> = {
  food: "Food & Pantry",
  household_supplies: "Household Supplies",
  personal_care: "Personal Care & Health",
};

export const visibilityLabels: Record<Visibility, string> = {
  private: "Private",
  shared: "Shared",
  household: "Household",
};

export function formatQuantity(value: number, unit: QuantityUnit): string {
  if (unit === "percentage") {
    return `${Math.round(value)}%`;
  }
  if (unit === "count") {
    return `${Math.round(value)} count`;
  }
  return `${Math.round(value)} ${unit}`;
}
