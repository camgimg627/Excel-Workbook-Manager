import { buildAssistantReply } from "@/lib/assistant";
import { createImportedRecipeDraft, createSeedDemoSession, instantiateCaptureDraft } from "@/lib/demo-data";
import {
  AssistantMessage,
  Category,
  DemoSessionState,
  InventoryInput,
  InventoryItem,
  InventoryUpdateInput,
  MemberRole,
  ProductUnitType,
  PurchaseEvent,
  QuantityUnit,
  RecipeImportDraft,
  ShoppingListEntry,
  ShoppingListInput,
} from "@/lib/domain";
import { getActiveHousehold, getActiveMember, getVisibleInventoryItems } from "@/lib/selectors";
import { clamp, createId } from "@/lib/utils";

function getRole(state: DemoSessionState): MemberRole {
  return getActiveMember(state).role;
}

function unitTypeFromQuantityUnit(unit: QuantityUnit): ProductUnitType {
  switch (unit) {
    case "ml":
      return "volume_ml";
    case "g":
      return "weight_g";
    case "percentage":
      return "percentage";
    default:
      return "count";
  }
}

function withUpdatedInventoryItem(
  state: DemoSessionState,
  inventoryItemId: string,
  updater: (item: InventoryItem) => InventoryItem
): DemoSessionState {
  return {
    ...state,
    inventoryItems: state.inventoryItems.map((entry) => (entry.id === inventoryItemId ? updater(entry) : entry)),
  };
}

function addPurchaseEvent(state: DemoSessionState, event: Omit<PurchaseEvent, "id">): DemoSessionState {
  return {
    ...state,
    purchaseEvents: [...state.purchaseEvents, { id: createId("purchase"), ...event }],
  };
}

function findOrCreateProduct(
  state: DemoSessionState,
  input: {
    name: string;
    category: Category;
    subcategory: string;
    quantityUnit: QuantityUnit;
    quantityFullValue: number;
    brand?: string;
  }
) {
  const existingProduct = state.products.find((entry) => entry.name.toLowerCase() === input.name.toLowerCase());

  if (existingProduct) {
    return { state, productId: existingProduct.id };
  }

  const nextProductId = createId("product");
  return {
    state: {
      ...state,
      products: [
        ...state.products,
        {
          id: nextProductId,
          name: input.name,
          brand: input.brand,
          category: input.category,
          subcategory: input.subcategory,
          unitType: unitTypeFromQuantityUnit(input.quantityUnit),
          defaultUnitSize: input.quantityFullValue,
          createdAt: new Date().toISOString(),
        },
      ],
    },
    productId: nextProductId,
  };
}

function findMatchingInventoryItem(
  state: DemoSessionState,
  input: {
    householdId: string;
    productId: string;
    spaceId: string;
    visibility: InventoryItem["visibility"];
    ownerUserId: string;
  }
) {
  return state.inventoryItems.find(
    (entry) =>
      entry.householdId === input.householdId &&
      entry.productId === input.productId &&
      entry.spaceId === input.spaceId &&
      entry.visibility === input.visibility &&
      (entry.visibility !== "private" || entry.ownerUserId === input.ownerUserId)
  );
}

function mergeQuantity(item: InventoryItem, quantityToAdd: number, quantityFullValue: number): InventoryItem {
  if (item.quantityUnit === "percentage") {
    return {
      ...item,
      quantityValue: clamp(Math.max(item.quantityValue, quantityToAdd), 0, 100),
      quantityFullValue: 100,
      updatedAt: new Date().toISOString(),
    };
  }

  const nextQuantity = item.quantityValue + quantityToAdd;
  return {
    ...item,
    quantityValue: nextQuantity,
    quantityFullValue: Math.max(item.quantityFullValue, quantityFullValue, nextQuantity),
    updatedAt: new Date().toISOString(),
  };
}

export function switchHousehold(state: DemoSessionState, householdId: string): DemoSessionState {
  const currentRole = getActiveMember(state).role;
  const matchingMember =
    state.householdMembers.find((entry) => entry.householdId === householdId && entry.role === currentRole) ??
    state.householdMembers.find((entry) => entry.householdId === householdId);

  if (!matchingMember) {
    return state;
  }

  return {
    ...state,
    selectedHouseholdId: householdId,
    selectedMemberId: matchingMember.id,
  };
}

export function switchMember(state: DemoSessionState, memberId: string): DemoSessionState {
  const member = state.householdMembers.find((entry) => entry.id === memberId);
  if (!member) {
    return state;
  }

  return {
    ...state,
    selectedHouseholdId: member.householdId,
    selectedMemberId: member.id,
  };
}

export function resetSession(): DemoSessionState {
  return createSeedDemoSession();
}

export function addInventory(state: DemoSessionState, input: InventoryInput): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const { state: withProduct, productId } = findOrCreateProduct(state, {
    name: input.name,
    category: input.category,
    subcategory: input.subcategory,
    quantityUnit: input.quantityUnit,
    quantityFullValue: input.quantityFullValue,
  });
  const matchingItem = findMatchingInventoryItem(withProduct, {
    householdId: household.id,
    productId,
    spaceId: input.spaceId,
    visibility: input.visibility,
    ownerUserId: member.userId,
  });

  if (matchingItem) {
    const mergedState = withUpdatedInventoryItem(withProduct, matchingItem.id, (entry) =>
      mergeQuantity(entry, input.quantityValue, input.quantityFullValue)
    );
    return input.retailer
      ? addPurchaseEvent(mergedState, {
          householdId: household.id,
          userId: member.userId,
          productId,
          quantityPurchased: input.quantityValue,
          quantityUnit: input.quantityUnit,
          unitPrice: input.unitPrice ?? 0,
          totalPrice: (input.unitPrice ?? 0) * Math.max(input.quantityValue, 1),
          retailer: input.retailer,
          source: "manual",
          purchasedAt: new Date().toISOString(),
        })
      : mergedState;
  }

  const nextState: DemoSessionState = {
    ...withProduct,
    inventoryItems: [
      ...withProduct.inventoryItems,
      {
        id: createId("inventory"),
        householdId: household.id,
        ownerUserId: member.userId,
        productId,
        spaceId: input.spaceId,
        visibility: input.visibility,
        quantityValue: input.quantityUnit === "percentage" ? clamp(input.quantityValue, 0, 100) : input.quantityValue,
        quantityUnit: input.quantityUnit,
        quantityFullValue: input.quantityUnit === "percentage" ? 100 : input.quantityFullValue,
        expiryDate: input.expiryDate,
        notes: input.notes,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    ],
  };

  return input.retailer
    ? addPurchaseEvent(nextState, {
        householdId: household.id,
        userId: member.userId,
        productId,
        quantityPurchased: input.quantityValue,
        quantityUnit: input.quantityUnit,
        unitPrice: input.unitPrice ?? 0,
        totalPrice: (input.unitPrice ?? 0) * Math.max(input.quantityValue, 1),
        retailer: input.retailer,
        source: "manual",
        purchasedAt: new Date().toISOString(),
      })
    : nextState;
}

export function updateInventory(state: DemoSessionState, inventoryItemId: string, input: InventoryUpdateInput): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  return withUpdatedInventoryItem(state, inventoryItemId, (entry) => ({
    ...entry,
    spaceId: input.spaceId,
    visibility: input.visibility,
    quantityValue: entry.quantityUnit === "percentage" ? clamp(input.quantityValue, 0, 100) : input.quantityValue,
    quantityFullValue: entry.quantityUnit === "percentage" ? 100 : input.quantityFullValue,
    expiryDate: input.expiryDate,
    notes: input.notes,
    updatedAt: new Date().toISOString(),
  }));
}

export function useInventory(state: DemoSessionState, inventoryItemId: string, quantityUsed: number): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  const member = getActiveMember(state);
  const item = state.inventoryItems.find((entry) => entry.id === inventoryItemId);
  if (!item) {
    return state;
  }

  const nextState = withUpdatedInventoryItem(state, inventoryItemId, (entry) => ({
    ...entry,
    quantityValue: clamp(entry.quantityValue - quantityUsed, 0, entry.quantityUnit === "percentage" ? 100 : Number.MAX_SAFE_INTEGER),
    updatedAt: new Date().toISOString(),
  }));

  return {
    ...nextState,
    consumptionEvents: [
      ...nextState.consumptionEvents,
      {
        id: createId("consume"),
        inventoryItemId,
        userId: member.userId,
        quantityUsed,
        source: "manual",
        consumedAt: new Date().toISOString(),
      },
    ],
  };
}

export function queueCapture(state: DemoSessionState, templateKey: "barcode" | "receipt" | "browser"): DemoSessionState {
  const household = getActiveHousehold(state);
  return {
    ...state,
    captureDrafts: [...state.captureDrafts, instantiateCaptureDraft(templateKey, household.id)],
  };
}

export function acceptCapture(state: DemoSessionState, draftId: string): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  const member = getActiveMember(state);
  const draft = state.captureDrafts.find((entry) => entry.id === draftId);
  if (!draft) {
    return state;
  }

  let nextState: DemoSessionState = {
    ...state,
    captureDrafts: state.captureDrafts.filter((entry) => entry.id !== draftId),
  };

  draft.items.forEach((entry) => {
    const productResult = findOrCreateProduct(nextState, {
      name: entry.productName,
      category: entry.category,
      subcategory: entry.subcategory,
      quantityUnit: entry.quantityUnit,
      quantityFullValue: entry.quantityFullValue,
      brand: entry.brand,
    });
    nextState = productResult.state;

    const matchingItem = findMatchingInventoryItem(nextState, {
      householdId: draft.householdId,
      productId: productResult.productId,
      spaceId: entry.spaceId,
      visibility: entry.visibility,
      ownerUserId: member.userId,
    });

    if (matchingItem) {
      nextState = withUpdatedInventoryItem(nextState, matchingItem.id, (item) =>
        mergeQuantity(item, entry.quantity, entry.quantityFullValue)
      );
    } else {
      nextState = {
        ...nextState,
        inventoryItems: [
          ...nextState.inventoryItems,
          {
            id: createId("inventory"),
            householdId: draft.householdId,
            ownerUserId: member.userId,
            productId: productResult.productId,
            spaceId: entry.spaceId,
            visibility: entry.visibility,
            quantityValue: entry.quantity,
            quantityUnit: entry.quantityUnit,
            quantityFullValue: entry.quantityUnit === "percentage" ? 100 : entry.quantityFullValue,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
          },
        ],
      };
    }

    nextState = addPurchaseEvent(nextState, {
      householdId: draft.householdId,
      userId: member.userId,
      productId: productResult.productId,
      quantityPurchased: entry.quantity,
      quantityUnit: entry.quantityUnit,
      unitPrice: 0,
      totalPrice: 0,
      retailer: draft.retailer,
      source: draft.method,
      purchasedAt: new Date().toISOString(),
    });
  });

  return nextState;
}

export function addShoppingItem(state: DemoSessionState, input: ShoppingListInput): DemoSessionState {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);

  return {
    ...state,
    shoppingListEntries: [
      ...state.shoppingListEntries,
      {
        id: createId("shopping"),
        householdId: household.id,
        productName: input.productName,
        category: input.category,
        quantity: input.quantity,
        quantityUnit: input.quantityUnit,
        spaceId: input.spaceId,
        note: input.note,
        source: "manual",
        createdBy: member.userId,
        createdAt: new Date().toISOString(),
      },
    ],
  };
}

export function purchaseShoppingItem(state: DemoSessionState, entry: ShoppingListEntry): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const { state: withProduct, productId } = findOrCreateProduct(state, {
    name: entry.productName,
    category: entry.category,
    subcategory: "shopping-list",
    quantityUnit: entry.quantityUnit,
    quantityFullValue: entry.quantity,
  });
  const visibleInventory = getVisibleInventoryItems(withProduct, household.id, member.id);
  const targetItem =
    (entry.inventoryItemId ? withProduct.inventoryItems.find((item) => item.id === entry.inventoryItemId) : undefined) ??
    visibleInventory.find((item) => item.productId === productId && item.visibility !== "private");

  let nextState = withProduct;

  if (targetItem) {
    nextState = withUpdatedInventoryItem(nextState, targetItem.id, (item) => mergeQuantity(item, entry.quantity, entry.quantity));
  } else {
    nextState = {
      ...nextState,
      inventoryItems: [
        ...nextState.inventoryItems,
        {
          id: createId("inventory"),
          householdId: household.id,
          ownerUserId: member.userId,
          productId,
          spaceId: entry.spaceId ?? nextState.spaces.find((space) => space.householdId === household.id)?.id ?? nextState.spaces[0].id,
          visibility: "shared",
          quantityValue: entry.quantity,
          quantityUnit: entry.quantityUnit,
          quantityFullValue: entry.quantityUnit === "percentage" ? 100 : entry.quantity,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        },
      ],
    };
  }

  nextState = addPurchaseEvent(nextState, {
    householdId: household.id,
    userId: member.userId,
    productId,
    quantityPurchased: entry.quantity,
    quantityUnit: entry.quantityUnit,
    unitPrice: 0,
    totalPrice: 0,
    retailer: "Shopping list run",
    source: "manual",
    purchasedAt: new Date().toISOString(),
  });

  return {
    ...nextState,
    shoppingListEntries: nextState.shoppingListEntries.filter((listEntry) => listEntry.id !== entry.id),
  };
}

export function cookRecipe(state: DemoSessionState, recipeId: string, servingsMade: number): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const recipe = state.recipes.find((entry) => entry.id === recipeId && entry.householdId === household.id);
  if (!recipe) {
    return state;
  }

  const multiplier = servingsMade / Math.max(recipe.servings, 1);
  let nextState = state;
  const recipeLogId = createId("recipe-log");

  nextState.recipeIngredients
    .filter((entry) => entry.recipeId === recipeId)
    .forEach((ingredient) => {
      const matchingItem = nextState.inventoryItems.find(
        (entry) => entry.householdId === household.id && entry.productId === ingredient.productId
      );
      if (!matchingItem) {
        return;
      }

      const quantityUsed = ingredient.quantity * multiplier;
      nextState = withUpdatedInventoryItem(nextState, matchingItem.id, (item) => ({
        ...item,
        quantityValue: clamp(item.quantityValue - quantityUsed, 0, item.quantityUnit === "percentage" ? 100 : Number.MAX_SAFE_INTEGER),
        updatedAt: new Date().toISOString(),
      }));
      nextState = {
        ...nextState,
        consumptionEvents: [
          ...nextState.consumptionEvents,
          {
            id: createId("consume"),
            inventoryItemId: matchingItem.id,
            userId: member.userId,
            quantityUsed,
            source: "recipe",
            recipeLogId,
            consumedAt: new Date().toISOString(),
          },
        ],
      };
    });

  return {
    ...nextState,
    recipeLogs: [
      ...nextState.recipeLogs,
      {
        id: recipeLogId,
        recipeId,
        householdId: household.id,
        cookedBy: member.userId,
        servingsMade,
        cookedAt: new Date().toISOString(),
        inventoryDeducted: true,
      },
    ],
  };
}

function importRecipeDraft(state: DemoSessionState, draft: RecipeImportDraft): DemoSessionState {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  let nextState = state;
  const recipeId = createId("recipe");

  nextState = {
    ...nextState,
    recipes: [
      ...nextState.recipes,
      {
        id: recipeId,
        householdId: household.id,
        createdBy: member.userId,
        name: draft.name,
        servings: draft.servings,
        sourceUrl: draft.sourceUrl,
        instructions: draft.instructions,
        createdAt: new Date().toISOString(),
      },
    ],
  };

  draft.ingredients.forEach((entry) => {
    const productResult = findOrCreateProduct(nextState, {
      name: entry.productName,
      category: entry.category,
      subcategory: entry.subcategory,
      quantityUnit: entry.unit,
      quantityFullValue: entry.quantity,
    });
    nextState = {
      ...productResult.state,
      recipeIngredients: [
        ...productResult.state.recipeIngredients,
        {
          id: createId("recipe-ingredient"),
          recipeId,
          productId: productResult.productId,
          quantity: entry.quantity,
          unit: entry.unit,
        },
      ],
    };
  });

  return nextState;
}

export function importRecipeFromUrl(state: DemoSessionState, url: string): DemoSessionState {
  if (getRole(state) === "contributor") {
    return state;
  }

  return importRecipeDraft(state, createImportedRecipeDraft(url));
}

export function askAssistant(state: DemoSessionState, prompt: string): DemoSessionState {
  const household = getActiveHousehold(state);
  const member = getActiveMember(state);
  const userMessage: AssistantMessage = {
    id: createId("assistant-user"),
    householdId: household.id,
    memberId: member.id,
    role: "user",
    content: prompt,
    createdAt: new Date().toISOString(),
  };
  const assistantMessage: AssistantMessage = {
    id: createId("assistant-reply"),
    householdId: household.id,
    memberId: member.id,
    role: "assistant",
    content: buildAssistantReply(state, prompt),
    createdAt: new Date().toISOString(),
  };

  return {
    ...state,
    assistantMessages: [...state.assistantMessages, userMessage, assistantMessage],
  };
}
