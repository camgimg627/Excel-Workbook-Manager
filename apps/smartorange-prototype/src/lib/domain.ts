export type HouseholdType = "individual" | "couple" | "family" | "multigenerational" | "roommates";
export type MemberRole = "owner" | "member" | "contributor";
export type Category = "food" | "household_supplies" | "personal_care";
export type QuantityUnit = "count" | "ml" | "g" | "percentage";
export type ProductUnitType = "count" | "volume_ml" | "weight_g" | "percentage";
export type Visibility = "private" | "shared" | "household";
export type CaptureMethod = "manual" | "barcode" | "receipt_scan" | "browser_extension";
export type ConsumptionSource = "manual" | "schedule" | "recipe";
export type PurchaseSource = "manual" | "barcode" | "receipt_scan" | "browser_extension";
export type ScheduleFrequency = "daily" | "weekly" | "custom";
export type ShoppingListSource = "manual" | "forecast" | "expiry";

export interface Household {
  id: string;
  name: string;
  type: HouseholdType;
  createdAt: string;
  settings: {
    alertLeadDays: number;
    replenishmentLeadDays: number;
  };
  summary: string;
}

export interface User {
  id: string;
  email: string;
  displayName: string;
  avatarUrl?: string;
  createdAt: string;
}

export interface HouseholdMember {
  id: string;
  householdId: string;
  userId: string;
  role: MemberRole;
  joinedAt: string;
  label: string;
}

export interface Product {
  id: string;
  name: string;
  brand?: string;
  barcode?: string;
  category: Category;
  subcategory: string;
  unitType: ProductUnitType;
  defaultUnitSize: number;
  imageUrl?: string;
  createdAt: string;
}

export interface Space {
  id: string;
  householdId: string;
  name: string;
  type: "pantry" | "bathroom" | "garage" | "bedroom" | "other";
  createdAt: string;
}

export interface InventoryItem {
  id: string;
  householdId: string;
  ownerUserId: string;
  productId: string;
  spaceId: string;
  visibility: Visibility;
  quantityValue: number;
  quantityUnit: QuantityUnit;
  quantityFullValue: number;
  expiryDate?: string;
  notes?: string;
  createdAt: string;
  updatedAt: string;
}

export interface PurchaseEvent {
  id: string;
  householdId: string;
  userId: string;
  productId: string;
  quantityPurchased: number;
  quantityUnit: QuantityUnit;
  unitPrice: number;
  totalPrice: number;
  retailer: string;
  source: PurchaseSource;
  purchasedAt: string;
  receiptUrl?: string;
}

export interface ConsumptionEvent {
  id: string;
  inventoryItemId: string;
  userId: string;
  quantityUsed: number;
  source: ConsumptionSource;
  recipeLogId?: string;
  consumedAt: string;
}

export interface ConsumptionSchedule {
  id: string;
  householdId: string;
  productId: string;
  quantityPerUse: number;
  frequency: ScheduleFrequency;
  frequencyDays: number[];
  notificationEnabled: boolean;
  lastPromptedAt?: string;
  createdAt: string;
}

export interface Recipe {
  id: string;
  householdId: string;
  createdBy: string;
  name: string;
  servings: number;
  sourceUrl?: string;
  instructions: string;
  createdAt: string;
}

export interface RecipeIngredient {
  id: string;
  recipeId: string;
  productId: string;
  quantity: number;
  unit: QuantityUnit;
}

export interface RecipeLog {
  id: string;
  recipeId: string;
  householdId: string;
  cookedBy: string;
  servingsMade: number;
  cookedAt: string;
  inventoryDeducted: boolean;
}

export interface ForecastSnapshot {
  inventoryItemId: string;
  householdId: string;
  productId: string;
  productName: string;
  category: Category;
  quantityUnit: QuantityUnit;
  quantityValue: number;
  quantityFullValue: number;
  dailyConsumptionRate: number;
  daysRemaining: number;
  replenishmentDate: string;
  alertDate: string;
  lowStock: boolean;
  confidence: "scheduled" | "historical" | "default";
}

export interface ShoppingListEntry {
  id: string;
  householdId: string;
  inventoryItemId?: string;
  productId?: string;
  productName: string;
  category: Category;
  quantity: number;
  quantityUnit: QuantityUnit;
  spaceId?: string;
  source: ShoppingListSource;
  note?: string;
  createdBy: string;
  createdAt: string;
}

export interface CaptureDraftItem {
  id: string;
  productName: string;
  brand?: string;
  category: Category;
  subcategory: string;
  quantity: number;
  quantityUnit: QuantityUnit;
  quantityFullValue: number;
  spaceId: string;
  visibility: Visibility;
  confidence: number;
}

export interface CaptureDraft {
  id: string;
  householdId: string;
  method: Exclude<CaptureMethod, "manual">;
  title: string;
  retailer: string;
  detectedAt: string;
  note: string;
  items: CaptureDraftItem[];
}

export interface AssistantMessage {
  id: string;
  householdId: string;
  memberId: string;
  role: "user" | "assistant";
  content: string;
  createdAt: string;
}

export interface NotificationItem {
  id: string;
  householdId: string;
  severity: "info" | "warning" | "urgent";
  title: string;
  body: string;
  routeHint?: string;
}

export interface DemoSessionState {
  households: Household[];
  users: User[];
  householdMembers: HouseholdMember[];
  products: Product[];
  spaces: Space[];
  inventoryItems: InventoryItem[];
  purchaseEvents: PurchaseEvent[];
  consumptionEvents: ConsumptionEvent[];
  consumptionSchedules: ConsumptionSchedule[];
  recipes: Recipe[];
  recipeIngredients: RecipeIngredient[];
  recipeLogs: RecipeLog[];
  shoppingListEntries: ShoppingListEntry[];
  captureDrafts: CaptureDraft[];
  assistantMessages: AssistantMessage[];
  selectedHouseholdId: string;
  selectedMemberId: string;
}

export interface InventoryInput {
  name: string;
  category: Category;
  subcategory: string;
  quantityValue: number;
  quantityUnit: QuantityUnit;
  quantityFullValue: number;
  spaceId: string;
  visibility: Visibility;
  expiryDate?: string;
  notes?: string;
  retailer?: string;
  unitPrice?: number;
}

export interface InventoryUpdateInput {
  spaceId: string;
  visibility: Visibility;
  quantityValue: number;
  quantityFullValue: number;
  expiryDate?: string;
  notes?: string;
}

export interface ShoppingListInput {
  productName: string;
  category: Category;
  quantity: number;
  quantityUnit: QuantityUnit;
  spaceId?: string;
  note?: string;
}

export interface RecipeImportDraft {
  name: string;
  servings: number;
  sourceUrl: string;
  instructions: string;
  ingredients: Array<{
    productName: string;
    category: Category;
    subcategory: string;
    quantity: number;
    unit: QuantityUnit;
  }>;
}
