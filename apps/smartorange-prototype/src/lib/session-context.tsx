"use client";

import * as React from "react";
import { createSeedDemoSession } from "@/lib/demo-data";
import { DemoSessionState, InventoryInput, InventoryUpdateInput, ShoppingListEntry, ShoppingListInput } from "@/lib/domain";
import {
  addInventory,
  addShoppingItem,
  askAssistant,
  cookRecipe,
  importRecipeFromUrl,
  purchaseShoppingItem,
  queueCapture,
  acceptCapture,
  resetSession,
  switchHousehold,
  switchMember,
  updateInventory,
  useInventory,
} from "@/lib/session-actions";
import {
  getActiveHousehold,
  getActiveMember,
  getActiveUser,
  getCurrentShoppingList,
  getDashboardMetrics,
  getForecastForActiveView,
  getNotifications,
  getRecentActivity,
  getRolePermissions,
  getSpaceSnapshots,
  getSpacesForHousehold,
  getVisibleInventoryItems,
} from "@/lib/selectors";

const STORAGE_KEY = "smartorange:prototype:session:v1";

interface DemoSessionContextValue {
  state: DemoSessionState;
  hydrated: boolean;
  activeHousehold: ReturnType<typeof getActiveHousehold>;
  activeMember: ReturnType<typeof getActiveMember>;
  activeUser: ReturnType<typeof getActiveUser>;
  visibleInventory: ReturnType<typeof getVisibleInventoryItems>;
  forecasts: ReturnType<typeof getForecastForActiveView>;
  shoppingList: ReturnType<typeof getCurrentShoppingList>;
  notifications: ReturnType<typeof getNotifications>;
  recentActivity: ReturnType<typeof getRecentActivity>;
  spaceSnapshots: ReturnType<typeof getSpaceSnapshots>;
  spaces: ReturnType<typeof getSpacesForHousehold>;
  metrics: ReturnType<typeof getDashboardMetrics>;
  permissions: ReturnType<typeof getRolePermissions>;
  actions: {
    switchHousehold: (householdId: string) => void;
    switchMember: (memberId: string) => void;
    reset: () => void;
    addInventory: (input: InventoryInput) => void;
    updateInventory: (inventoryItemId: string, input: InventoryUpdateInput) => void;
    useInventory: (inventoryItemId: string, quantityUsed: number) => void;
    queueCapture: (templateKey: "barcode" | "receipt" | "browser") => void;
    acceptCapture: (draftId: string) => void;
    addShoppingItem: (input: ShoppingListInput) => void;
    purchaseShoppingItem: (entry: ShoppingListEntry) => void;
    cookRecipe: (recipeId: string, servingsMade: number) => void;
    importRecipeFromUrl: (url: string) => void;
    askAssistant: (prompt: string) => void;
  };
}

const DemoSessionContext = React.createContext<DemoSessionContextValue | null>(null);

export function DemoSessionProvider({ children }: { children: React.ReactNode }) {
  const [state, setState] = React.useState<DemoSessionState>(createSeedDemoSession);
  const [hydrated, setHydrated] = React.useState(false);

  React.useEffect(() => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored) {
        setState(JSON.parse(stored) as DemoSessionState);
      }
    } catch {
      setState(createSeedDemoSession());
    } finally {
      setHydrated(true);
    }
  }, []);

  React.useEffect(() => {
    if (!hydrated) {
      return;
    }

    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state, hydrated]);

  const activeHousehold = getActiveHousehold(state);
  const activeMember = getActiveMember(state);
  const activeUser = getActiveUser(state);
  const visibleInventory = getVisibleInventoryItems(state, activeHousehold.id, activeMember.id);
  const forecasts = getForecastForActiveView(state);
  const shoppingList = getCurrentShoppingList(state);
  const notifications = getNotifications(state);
  const recentActivity = getRecentActivity(state);
  const spaceSnapshots = getSpaceSnapshots(state);
  const spaces = getSpacesForHousehold(state, activeHousehold.id);
  const metrics = getDashboardMetrics(state);
  const permissions = getRolePermissions(activeMember.role);

  const value: DemoSessionContextValue = {
    state,
    hydrated,
    activeHousehold,
    activeMember,
    activeUser,
    visibleInventory,
    forecasts,
    shoppingList,
    notifications,
    recentActivity,
    spaceSnapshots,
    spaces,
    metrics,
    permissions,
    actions: {
      switchHousehold: (householdId) => setState((current) => switchHousehold(current, householdId)),
      switchMember: (memberId) => setState((current) => switchMember(current, memberId)),
      reset: () => setState(resetSession()),
      addInventory: (input) => setState((current) => addInventory(current, input)),
      updateInventory: (inventoryItemId, input) => setState((current) => updateInventory(current, inventoryItemId, input)),
      useInventory: (inventoryItemId, quantityUsed) => setState((current) => useInventory(current, inventoryItemId, quantityUsed)),
      queueCapture: (templateKey) => setState((current) => queueCapture(current, templateKey)),
      acceptCapture: (draftId) => setState((current) => acceptCapture(current, draftId)),
      addShoppingItem: (input) => setState((current) => addShoppingItem(current, input)),
      purchaseShoppingItem: (entry) => setState((current) => purchaseShoppingItem(current, entry)),
      cookRecipe: (recipeId, servingsMade) => setState((current) => cookRecipe(current, recipeId, servingsMade)),
      importRecipeFromUrl: (url) => setState((current) => importRecipeFromUrl(current, url)),
      askAssistant: (prompt) => setState((current) => askAssistant(current, prompt)),
    },
  };

  return <DemoSessionContext.Provider value={value}>{children}</DemoSessionContext.Provider>;
}

export function useDemoSession() {
  const context = React.useContext(DemoSessionContext);
  if (!context) {
    throw new Error("useDemoSession must be used within DemoSessionProvider.");
  }
  return context;
}
