import {
  AssistantMessage,
  CaptureDraft,
  CaptureDraftItem,
  Category,
  DemoSessionState,
  Household,
  HouseholdMember,
  InventoryItem,
  Product,
  PurchaseEvent,
  QuantityUnit,
  Recipe,
  RecipeImportDraft,
  RecipeIngredient,
  RecipeLog,
  ShoppingListEntry,
  Space,
  User,
  ConsumptionEvent,
  ConsumptionSchedule,
  Visibility,
} from "@/lib/domain";
import { createId } from "@/lib/utils";

const BASE_CREATED_AT = "2026-04-01T12:00:00.000Z";

function user(id: string, displayName: string, email: string): User {
  return {
    id,
    displayName,
    email,
    createdAt: BASE_CREATED_AT,
  };
}

function household(
  id: string,
  name: string,
  type: Household["type"],
  summary: string,
  settings: Household["settings"]
): Household {
  return {
    id,
    name,
    type,
    createdAt: BASE_CREATED_AT,
    summary,
    settings,
  };
}

function member(
  id: string,
  householdId: string,
  userId: string,
  role: HouseholdMember["role"],
  label: string
): HouseholdMember {
  return {
    id,
    householdId,
    userId,
    role,
    joinedAt: BASE_CREATED_AT,
    label,
  };
}

function product(
  id: string,
  name: string,
  category: Category,
  subcategory: string,
  unitType: Product["unitType"],
  defaultUnitSize: number,
  brand?: string
): Product {
  return {
    id,
    name,
    category,
    subcategory,
    unitType,
    defaultUnitSize,
    brand,
    createdAt: BASE_CREATED_AT,
  };
}

function space(
  id: string,
  householdId: string,
  name: string,
  type: Space["type"]
): Space {
  return {
    id,
    householdId,
    name,
    type,
    createdAt: BASE_CREATED_AT,
  };
}

function inventoryItem(input: {
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
  updatedAt?: string;
}): InventoryItem {
  return {
    createdAt: BASE_CREATED_AT,
    updatedAt: input.updatedAt ?? "2026-04-06T15:00:00.000Z",
    ...input,
  };
}

function purchaseEvent(input: Omit<PurchaseEvent, "id"> & { id?: string }): PurchaseEvent {
  return {
    id: input.id ?? createId("purchase"),
    ...input,
  };
}

function consumptionSchedule(input: ConsumptionSchedule): ConsumptionSchedule {
  return input;
}

function consumptionEvent(input: ConsumptionEvent): ConsumptionEvent {
  return input;
}

function recipe(input: Recipe): Recipe {
  return input;
}

function ingredient(input: RecipeIngredient): RecipeIngredient {
  return input;
}

function recipeLog(input: RecipeLog): RecipeLog {
  return input;
}

function shoppingEntry(input: ShoppingListEntry): ShoppingListEntry {
  return input;
}

function assistantMessage(input: AssistantMessage): AssistantMessage {
  return input;
}

function draftItem(input: CaptureDraftItem): CaptureDraftItem {
  return input;
}

function captureDraft(input: CaptureDraft): CaptureDraft {
  return input;
}

export const households: Household[] = [
  household(
    "household-ava",
    "Ava's Apartment",
    "individual",
    "Solo household focused on reducing waste, catching expiry risk, and keeping pantry basics visible.",
    { alertLeadDays: 4, replenishmentLeadDays: 6 }
  ),
  household(
    "household-rivera",
    "Rivera Family",
    "family",
    "Two adults plus a teen coordinating shared staples, schedule-driven meals, and shopping handoffs.",
    { alertLeadDays: 5, replenishmentLeadDays: 7 }
  ),
  household(
    "household-maple",
    "Maple House Roommates",
    "roommates",
    "Shared consumables plus private items, with just enough visibility to avoid awkward duplicate buys.",
    { alertLeadDays: 4, replenishmentLeadDays: 5 }
  ),
];

export const users: User[] = [
  user("user-ava-owner", "Ava", "ava@smartorange.demo"),
  user("user-ava-member", "Lena", "lena@smartorange.demo"),
  user("user-ava-contributor", "Chris", "chris@smartorange.demo"),
  user("user-rivera-owner", "Alex Rivera", "alex@smartorange.demo"),
  user("user-rivera-member", "Morgan Rivera", "morgan@smartorange.demo"),
  user("user-rivera-contributor", "Mia Rivera", "mia@smartorange.demo"),
  user("user-maple-owner", "Noah", "noah@smartorange.demo"),
  user("user-maple-member", "Priya", "priya@smartorange.demo"),
  user("user-maple-contributor", "Eli", "eli@smartorange.demo"),
];

export const householdMembers: HouseholdMember[] = [
  member("member-ava-owner", "household-ava", "user-ava-owner", "owner", "Primary household owner"),
  member("member-ava-member", "household-ava", "user-ava-member", "member", "Trusted helper view"),
  member("member-ava-contributor", "household-ava", "user-ava-contributor", "contributor", "Errand helper"),
  member("member-rivera-owner", "household-rivera", "user-rivera-owner", "owner", "Parent admin"),
  member("member-rivera-member", "household-rivera", "user-rivera-member", "member", "Shared household manager"),
  member("member-rivera-contributor", "household-rivera", "user-rivera-contributor", "contributor", "Teen runner"),
  member("member-maple-owner", "household-maple", "user-maple-owner", "owner", "Lease holder"),
  member("member-maple-member", "household-maple", "user-maple-member", "member", "Equal roommate"),
  member("member-maple-contributor", "household-maple", "user-maple-contributor", "contributor", "Weekend helper"),
];

export const products: Product[] = [
  product("product-olive-oil", "Olive Oil", "food", "pantry", "volume_ml", 500, "Solea"),
  product("product-spaghetti", "Spaghetti", "food", "pasta", "weight_g", 500, "Bronze Cut"),
  product("product-coffee", "Coffee Beans", "food", "coffee", "weight_g", 340, "Crescent Roasters"),
  product("product-spinach", "Baby Spinach", "food", "produce", "weight_g", 300, "Farm Lane"),
  product("product-shampoo", "Shampoo", "personal_care", "haircare", "percentage", 100, "Citrus Bloom"),
  product("product-dish-soap", "Dish Soap", "household_supplies", "cleaning", "percentage", 100, "Spark"),
  product("product-toilet-paper", "Toilet Paper", "household_supplies", "paper", "count", 12, "CloudSoft"),
  product("product-paper-towels", "Paper Towels", "household_supplies", "paper", "count", 6, "CloudSoft"),
  product("product-milk", "Milk", "food", "dairy", "volume_ml", 2000, "North Dairy"),
  product("product-pasta-sauce", "Pasta Sauce", "food", "sauce", "volume_ml", 650, "Sunday Jar"),
  product("product-detergent", "Laundry Detergent", "household_supplies", "laundry", "percentage", 100, "Bright Wash"),
  product("product-allergy-tabs", "Allergy Relief", "personal_care", "health", "count", 24, "ClearDay"),
  product("product-ramen", "Ramen Pack", "food", "pantry", "count", 6, "Night Owl"),
  product("product-oat-milk", "Oat Milk", "food", "dairy", "volume_ml", 1900, "Oak Carton"),
  product("product-toothpaste", "Toothpaste", "personal_care", "dental", "percentage", 100, "Mint Loop"),
  product("product-frozen-veg", "Frozen Stir-Fry Mix", "food", "freezer", "weight_g", 400, "Market Bowl"),
];

export const spaces: Space[] = [
  space("space-ava-pantry", "household-ava", "Pantry Shelf", "pantry"),
  space("space-ava-bathroom", "household-ava", "Bathroom Cabinet", "bathroom"),
  space("space-ava-utility", "household-ava", "Utility Closet", "other"),
  space("space-rivera-pantry", "household-rivera", "Main Pantry", "pantry"),
  space("space-rivera-bathroom", "household-rivera", "Kids Bathroom", "bathroom"),
  space("space-rivera-garage", "household-rivera", "Garage Shelf", "garage"),
  space("space-maple-pantry", "household-maple", "Shared Pantry", "pantry"),
  space("space-maple-bathroom", "household-maple", "Hall Bathroom", "bathroom"),
  space("space-maple-freezer", "household-maple", "Freezer Drawer", "other"),
];

export const inventoryItems: InventoryItem[] = [
  inventoryItem({
    id: "item-ava-olive-oil",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-olive-oil",
    spaceId: "space-ava-pantry",
    visibility: "household",
    quantityValue: 180,
    quantityUnit: "ml",
    quantityFullValue: 500,
    notes: "Mostly used for pasta nights.",
  }),
  inventoryItem({
    id: "item-ava-spaghetti",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-spaghetti",
    spaceId: "space-ava-pantry",
    visibility: "household",
    quantityValue: 260,
    quantityUnit: "g",
    quantityFullValue: 500,
  }),
  inventoryItem({
    id: "item-ava-coffee",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-coffee",
    spaceId: "space-ava-pantry",
    visibility: "household",
    quantityValue: 120,
    quantityUnit: "g",
    quantityFullValue: 340,
  }),
  inventoryItem({
    id: "item-ava-spinach",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-spinach",
    spaceId: "space-ava-pantry",
    visibility: "household",
    quantityValue: 110,
    quantityUnit: "g",
    quantityFullValue: 300,
    expiryDate: "2026-04-09",
  }),
  inventoryItem({
    id: "item-ava-shampoo",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-shampoo",
    spaceId: "space-ava-bathroom",
    visibility: "private",
    quantityValue: 42,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-ava-dish-soap",
    householdId: "household-ava",
    ownerUserId: "user-ava-owner",
    productId: "product-dish-soap",
    spaceId: "space-ava-utility",
    visibility: "household",
    quantityValue: 65,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-rivera-toilet-paper",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-owner",
    productId: "product-toilet-paper",
    spaceId: "space-rivera-garage",
    visibility: "shared",
    quantityValue: 3,
    quantityUnit: "count",
    quantityFullValue: 12,
  }),
  inventoryItem({
    id: "item-rivera-paper-towels",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-owner",
    productId: "product-paper-towels",
    spaceId: "space-rivera-garage",
    visibility: "shared",
    quantityValue: 1,
    quantityUnit: "count",
    quantityFullValue: 6,
  }),
  inventoryItem({
    id: "item-rivera-milk",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-member",
    productId: "product-milk",
    spaceId: "space-rivera-pantry",
    visibility: "shared",
    quantityValue: 650,
    quantityUnit: "ml",
    quantityFullValue: 2000,
    expiryDate: "2026-04-10",
  }),
  inventoryItem({
    id: "item-rivera-pasta-sauce",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-member",
    productId: "product-pasta-sauce",
    spaceId: "space-rivera-pantry",
    visibility: "shared",
    quantityValue: 240,
    quantityUnit: "ml",
    quantityFullValue: 650,
  }),
  inventoryItem({
    id: "item-rivera-spaghetti",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-member",
    productId: "product-spaghetti",
    spaceId: "space-rivera-pantry",
    visibility: "shared",
    quantityValue: 320,
    quantityUnit: "g",
    quantityFullValue: 500,
  }),
  inventoryItem({
    id: "item-rivera-olive-oil",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-owner",
    productId: "product-olive-oil",
    spaceId: "space-rivera-pantry",
    visibility: "shared",
    quantityValue: 90,
    quantityUnit: "ml",
    quantityFullValue: 500,
  }),
  inventoryItem({
    id: "item-rivera-shampoo",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-owner",
    productId: "product-shampoo",
    spaceId: "space-rivera-bathroom",
    visibility: "household",
    quantityValue: 58,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-rivera-allergy",
    householdId: "household-rivera",
    ownerUserId: "user-rivera-member",
    productId: "product-allergy-tabs",
    spaceId: "space-rivera-bathroom",
    visibility: "private",
    quantityValue: 5,
    quantityUnit: "count",
    quantityFullValue: 24,
  }),
  inventoryItem({
    id: "item-maple-dish-soap",
    householdId: "household-maple",
    ownerUserId: "user-maple-owner",
    productId: "product-dish-soap",
    spaceId: "space-maple-pantry",
    visibility: "shared",
    quantityValue: 18,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-maple-ramen",
    householdId: "household-maple",
    ownerUserId: "user-maple-owner",
    productId: "product-ramen",
    spaceId: "space-maple-pantry",
    visibility: "private",
    quantityValue: 2,
    quantityUnit: "count",
    quantityFullValue: 6,
  }),
  inventoryItem({
    id: "item-maple-oat-milk",
    householdId: "household-maple",
    ownerUserId: "user-maple-member",
    productId: "product-oat-milk",
    spaceId: "space-maple-pantry",
    visibility: "private",
    quantityValue: 800,
    quantityUnit: "ml",
    quantityFullValue: 1900,
    expiryDate: "2026-04-12",
  }),
  inventoryItem({
    id: "item-maple-toothpaste",
    householdId: "household-maple",
    ownerUserId: "user-maple-owner",
    productId: "product-toothpaste",
    spaceId: "space-maple-bathroom",
    visibility: "private",
    quantityValue: 35,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-maple-toilet-paper",
    householdId: "household-maple",
    ownerUserId: "user-maple-owner",
    productId: "product-toilet-paper",
    spaceId: "space-maple-pantry",
    visibility: "shared",
    quantityValue: 8,
    quantityUnit: "count",
    quantityFullValue: 12,
  }),
  inventoryItem({
    id: "item-maple-coffee",
    householdId: "household-maple",
    ownerUserId: "user-maple-member",
    productId: "product-coffee",
    spaceId: "space-maple-pantry",
    visibility: "shared",
    quantityValue: 80,
    quantityUnit: "g",
    quantityFullValue: 340,
  }),
  inventoryItem({
    id: "item-maple-shampoo",
    householdId: "household-maple",
    ownerUserId: "user-maple-member",
    productId: "product-shampoo",
    spaceId: "space-maple-bathroom",
    visibility: "private",
    quantityValue: 70,
    quantityUnit: "percentage",
    quantityFullValue: 100,
  }),
  inventoryItem({
    id: "item-maple-frozen-veg",
    householdId: "household-maple",
    ownerUserId: "user-maple-member",
    productId: "product-frozen-veg",
    spaceId: "space-maple-freezer",
    visibility: "shared",
    quantityValue: 220,
    quantityUnit: "g",
    quantityFullValue: 400,
  }),
];

export const purchaseEvents: PurchaseEvent[] = [
  purchaseEvent({
    householdId: "household-ava",
    userId: "user-ava-owner",
    productId: "product-coffee",
    quantityPurchased: 340,
    quantityUnit: "g",
    unitPrice: 12.99,
    totalPrice: 12.99,
    retailer: "Local Roastery",
    source: "manual",
    purchasedAt: "2026-04-03T14:05:00.000Z",
  }),
  purchaseEvent({
    householdId: "household-rivera",
    userId: "user-rivera-owner",
    productId: "product-paper-towels",
    quantityPurchased: 6,
    quantityUnit: "count",
    unitPrice: 10.5,
    totalPrice: 10.5,
    retailer: "Costco",
    source: "manual",
    purchasedAt: "2026-03-29T18:15:00.000Z",
  }),
  purchaseEvent({
    householdId: "household-rivera",
    userId: "user-rivera-member",
    productId: "product-milk",
    quantityPurchased: 2000,
    quantityUnit: "ml",
    unitPrice: 4.89,
    totalPrice: 4.89,
    retailer: "FreshMart",
    source: "manual",
    purchasedAt: "2026-04-05T17:10:00.000Z",
  }),
  purchaseEvent({
    householdId: "household-maple",
    userId: "user-maple-member",
    productId: "product-oat-milk",
    quantityPurchased: 1900,
    quantityUnit: "ml",
    unitPrice: 5.25,
    totalPrice: 5.25,
    retailer: "Night Grocer",
    source: "manual",
    purchasedAt: "2026-04-04T20:25:00.000Z",
  }),
];

export const consumptionSchedules: ConsumptionSchedule[] = [
  consumptionSchedule({
    id: "schedule-ava-coffee",
    householdId: "household-ava",
    productId: "product-coffee",
    quantityPerUse: 18,
    frequency: "daily",
    frequencyDays: [],
    notificationEnabled: true,
    lastPromptedAt: "2026-04-06T08:15:00.000Z",
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-ava-olive",
    householdId: "household-ava",
    productId: "product-olive-oil",
    quantityPerUse: 15,
    frequency: "custom",
    frequencyDays: [2, 5],
    notificationEnabled: false,
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-rivera-milk",
    householdId: "household-rivera",
    productId: "product-milk",
    quantityPerUse: 180,
    frequency: "daily",
    frequencyDays: [],
    notificationEnabled: true,
    lastPromptedAt: "2026-04-07T07:10:00.000Z",
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-rivera-toilet-paper",
    householdId: "household-rivera",
    productId: "product-toilet-paper",
    quantityPerUse: 1,
    frequency: "custom",
    frequencyDays: [0, 2, 4, 6],
    notificationEnabled: false,
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-rivera-paper-towels",
    householdId: "household-rivera",
    productId: "product-paper-towels",
    quantityPerUse: 1,
    frequency: "weekly",
    frequencyDays: [],
    notificationEnabled: false,
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-rivera-spaghetti",
    householdId: "household-rivera",
    productId: "product-spaghetti",
    quantityPerUse: 320,
    frequency: "weekly",
    frequencyDays: [],
    notificationEnabled: true,
    lastPromptedAt: "2026-04-01T17:00:00.000Z",
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-maple-dish",
    householdId: "household-maple",
    productId: "product-dish-soap",
    quantityPerUse: 7,
    frequency: "custom",
    frequencyDays: [1, 4],
    notificationEnabled: false,
    createdAt: BASE_CREATED_AT,
  }),
  consumptionSchedule({
    id: "schedule-maple-coffee",
    householdId: "household-maple",
    productId: "product-coffee",
    quantityPerUse: 22,
    frequency: "daily",
    frequencyDays: [],
    notificationEnabled: true,
    createdAt: BASE_CREATED_AT,
  }),
];

export const consumptionEvents: ConsumptionEvent[] = [
  consumptionEvent({
    id: "consume-1",
    inventoryItemId: "item-rivera-milk",
    userId: "user-rivera-contributor",
    quantityUsed: 250,
    source: "manual",
    consumedAt: "2026-04-06T08:30:00.000Z",
  }),
  consumptionEvent({
    id: "consume-2",
    inventoryItemId: "item-maple-dish-soap",
    userId: "user-maple-owner",
    quantityUsed: 8,
    source: "manual",
    consumedAt: "2026-04-05T19:20:00.000Z",
  }),
];

export const recipes: Recipe[] = [
  recipe({
    id: "recipe-ava-pasta",
    householdId: "household-ava",
    createdBy: "user-ava-owner",
    name: "Pantry Pasta Bowl",
    servings: 2,
    sourceUrl: "https://smartorange.demo/recipes/pantry-pasta",
    instructions: "Boil pasta, toss with olive oil, wilt spinach, and finish with pepper.",
    createdAt: BASE_CREATED_AT,
  }),
  recipe({
    id: "recipe-rivera-spaghetti",
    householdId: "household-rivera",
    createdBy: "user-rivera-member",
    name: "Spaghetti Tuesday",
    servings: 4,
    sourceUrl: "https://smartorange.demo/recipes/spaghetti-tuesday",
    instructions: "Cook pasta, warm sauce, saute spinach in olive oil, then combine.",
    createdAt: BASE_CREATED_AT,
  }),
  recipe({
    id: "recipe-maple-stirfry",
    householdId: "household-maple",
    createdBy: "user-maple-member",
    name: "Freezer Stir-Fry Rescue",
    servings: 2,
    sourceUrl: "https://smartorange.demo/recipes/freezer-stirfry",
    instructions: "Heat frozen veg, boil ramen, then combine with pantry sauce packet.",
    createdAt: BASE_CREATED_AT,
  }),
];

export const recipeIngredients: RecipeIngredient[] = [
  ingredient({ id: "ingredient-ava-1", recipeId: "recipe-ava-pasta", productId: "product-spaghetti", quantity: 180, unit: "g" }),
  ingredient({ id: "ingredient-ava-2", recipeId: "recipe-ava-pasta", productId: "product-olive-oil", quantity: 12, unit: "ml" }),
  ingredient({ id: "ingredient-ava-3", recipeId: "recipe-ava-pasta", productId: "product-spinach", quantity: 80, unit: "g" }),
  ingredient({
    id: "ingredient-rivera-1",
    recipeId: "recipe-rivera-spaghetti",
    productId: "product-spaghetti",
    quantity: 320,
    unit: "g",
  }),
  ingredient({
    id: "ingredient-rivera-2",
    recipeId: "recipe-rivera-spaghetti",
    productId: "product-pasta-sauce",
    quantity: 450,
    unit: "ml",
  }),
  ingredient({
    id: "ingredient-rivera-3",
    recipeId: "recipe-rivera-spaghetti",
    productId: "product-olive-oil",
    quantity: 15,
    unit: "ml",
  }),
  ingredient({
    id: "ingredient-maple-1",
    recipeId: "recipe-maple-stirfry",
    productId: "product-ramen",
    quantity: 2,
    unit: "count",
  }),
  ingredient({
    id: "ingredient-maple-2",
    recipeId: "recipe-maple-stirfry",
    productId: "product-frozen-veg",
    quantity: 180,
    unit: "g",
  }),
];

export const recipeLogs: RecipeLog[] = [
  recipeLog({
    id: "recipe-log-rivera-1",
    recipeId: "recipe-rivera-spaghetti",
    householdId: "household-rivera",
    cookedBy: "user-rivera-member",
    servingsMade: 4,
    cookedAt: "2026-04-01T18:40:00.000Z",
    inventoryDeducted: true,
  }),
];

export const shoppingListEntries: ShoppingListEntry[] = [
  shoppingEntry({
    id: "shopping-rivera-1",
    householdId: "household-rivera",
    productName: "Fruit Snacks",
    category: "food",
    quantity: 2,
    quantityUnit: "count",
    note: "Lunchbox refill before Friday.",
    createdBy: "user-rivera-contributor",
    createdAt: "2026-04-07T15:40:00.000Z",
    source: "manual",
  }),
  shoppingEntry({
    id: "shopping-maple-1",
    householdId: "household-maple",
    productName: "Sparkling Water",
    category: "food",
    quantity: 1,
    quantityUnit: "count",
    note: "Shared movie night stock-up.",
    createdBy: "user-maple-contributor",
    createdAt: "2026-04-06T18:15:00.000Z",
    source: "manual",
  }),
];

export const assistantMessages: AssistantMessage[] = [
  assistantMessage({
    id: "assistant-ava",
    householdId: "household-ava",
    memberId: "member-ava-owner",
    role: "assistant",
    content: "I can summarize low stock, expiry risk, last purchases, and whether your pantry covers tonight's plan.",
    createdAt: "2026-04-07T08:00:00.000Z",
  }),
  assistantMessage({
    id: "assistant-rivera",
    householdId: "household-rivera",
    memberId: "member-rivera-owner",
    role: "assistant",
    content: "Ask me about shared staples, shopping list gaps, or whether Tuesday dinner still fits what the family has left.",
    createdAt: "2026-04-07T08:00:00.000Z",
  }),
  assistantMessage({
    id: "assistant-maple",
    householdId: "household-maple",
    memberId: "member-maple-owner",
    role: "assistant",
    content: "I keep shared goods visible while respecting roommate privacy. Try asking what is low in shared spaces this week.",
    createdAt: "2026-04-07T08:00:00.000Z",
  }),
];

export const captureTemplates: Record<
  string,
  {
    method: CaptureDraft["method"];
    title: string;
    retailer: string;
    note: string;
    items: Array<{
      productName: string;
      brand?: string;
      category: Category;
      subcategory: string;
      quantity: number;
      quantityUnit: QuantityUnit;
      quantityFullValue: number;
      defaultVisibility: Visibility;
      confidence: number;
    }>;
  }
> = {
  barcode: {
    method: "barcode",
    title: "Barcode scan ready",
    retailer: "Kitchen counter",
    note: "Single-item capture from a barcode scan with catalog autofill.",
    items: [
      {
        productName: "Olive Oil",
        brand: "Solea",
        category: "food",
        subcategory: "pantry",
        quantity: 500,
        quantityUnit: "ml",
        quantityFullValue: 500,
        defaultVisibility: "shared",
        confidence: 0.98,
      },
    ],
  },
  receipt: {
    method: "receipt_scan",
    title: "Receipt parsed into inventory candidates",
    retailer: "FreshMart",
    note: "AI-style receipt parsing turns one photo into structured inventory items.",
    items: [
      {
        productName: "Milk",
        brand: "North Dairy",
        category: "food",
        subcategory: "dairy",
        quantity: 2000,
        quantityUnit: "ml",
        quantityFullValue: 2000,
        defaultVisibility: "shared",
        confidence: 0.93,
      },
      {
        productName: "Paper Towels",
        brand: "CloudSoft",
        category: "household_supplies",
        subcategory: "paper",
        quantity: 6,
        quantityUnit: "count",
        quantityFullValue: 6,
        defaultVisibility: "shared",
        confidence: 0.96,
      },
    ],
  },
  browser: {
    method: "browser_extension",
    title: "Retail cart captured before delivery",
    retailer: "Amazon",
    note: "Passive browser capture suggests inventory updates from online purchases.",
    items: [
      {
        productName: "Toilet Paper",
        brand: "CloudSoft",
        category: "household_supplies",
        subcategory: "paper",
        quantity: 12,
        quantityUnit: "count",
        quantityFullValue: 12,
        defaultVisibility: "shared",
        confidence: 0.97,
      },
      {
        productName: "Shampoo",
        brand: "Citrus Bloom",
        category: "personal_care",
        subcategory: "haircare",
        quantity: 100,
        quantityUnit: "percentage",
        quantityFullValue: 100,
        defaultVisibility: "private",
        confidence: 0.9,
      },
    ],
  },
};

export function instantiateCaptureDraft(templateKey: keyof typeof captureTemplates, householdId: string): CaptureDraft {
  const template = captureTemplates[templateKey];
  const householdSpaces = spaces.filter((entry) => entry.householdId === householdId);

  function resolveSpaceId(category: Category): string {
    if (category === "food") {
      return (
        householdSpaces.find((entry) => entry.type === "pantry")?.id ??
        householdSpaces.find((entry) => entry.type === "other")?.id ??
        householdSpaces[0]?.id ??
        spaces[0].id
      );
    }

    if (category === "household_supplies") {
      return (
        householdSpaces.find((entry) => entry.type === "garage")?.id ??
        householdSpaces.find((entry) => entry.type === "other")?.id ??
        householdSpaces.find((entry) => entry.type === "pantry")?.id ??
        householdSpaces[0]?.id ??
        spaces[0].id
      );
    }

    return (
      householdSpaces.find((entry) => entry.type === "bathroom")?.id ??
      householdSpaces.find((entry) => entry.type === "bedroom")?.id ??
      householdSpaces[0]?.id ??
      spaces[0].id
    );
  }

  return captureDraft({
    id: createId("capture"),
    householdId,
    method: template.method,
    title: template.title,
    retailer: template.retailer,
    detectedAt: "2026-04-07T16:20:00.000Z",
    note: template.note,
    items: template.items.map((entry) =>
      draftItem({
        id: createId("capture-item"),
        productName: entry.productName,
        brand: entry.brand,
        category: entry.category,
        subcategory: entry.subcategory,
        quantity: entry.quantity,
        quantityUnit: entry.quantityUnit,
        quantityFullValue: entry.quantityFullValue,
        visibility: entry.defaultVisibility,
        confidence: entry.confidence,
        spaceId: resolveSpaceId(entry.category),
      })
    ),
  });
}

export function createImportedRecipeDraft(url: string): RecipeImportDraft {
  return {
    name: "Weeknight Pantry Chili",
    servings: 4,
    sourceUrl: url,
    instructions: "Simulated recipe import: toast aromatics, simmer sauce, and batch for leftovers.",
    ingredients: [
      {
        productName: "Olive Oil",
        category: "food",
        subcategory: "pantry",
        quantity: 18,
        unit: "ml",
      },
      {
        productName: "Pasta Sauce",
        category: "food",
        subcategory: "sauce",
        quantity: 500,
        unit: "ml",
      },
      {
        productName: "Frozen Stir-Fry Mix",
        category: "food",
        subcategory: "freezer",
        quantity: 220,
        unit: "g",
      },
    ],
  };
}

export function createSeedDemoSession(): DemoSessionState {
  return {
    households,
    users,
    householdMembers,
    products,
    spaces,
    inventoryItems,
    purchaseEvents,
    consumptionEvents,
    consumptionSchedules,
    recipes,
    recipeIngredients,
    recipeLogs,
    shoppingListEntries,
    captureDrafts: [instantiateCaptureDraft("receipt", "household-rivera")],
    assistantMessages,
    selectedHouseholdId: "household-rivera",
    selectedMemberId: "member-rivera-owner",
  };
}
