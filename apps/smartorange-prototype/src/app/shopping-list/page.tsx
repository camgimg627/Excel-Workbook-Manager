"use client";

import * as React from "react";
import { Category, QuantityUnit } from "@/lib/domain";
import { categoryLabels, formatQuantity } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";
import { EmptyState, PageHeader, PermissionNotice, SectionTitle, Surface } from "@/components/ui";

export default function ShoppingListPage() {
  const { permissions, shoppingList, spaces, state, actions } = useDemoSession();
  const [form, setForm] = React.useState<{
    productName: string;
    category: Category;
    quantity: number;
    quantityUnit: QuantityUnit;
    spaceId: string;
    note: string;
  }>({
    productName: "",
    category: "food",
    quantity: 1,
    quantityUnit: "count",
    spaceId: spaces[0]?.id ?? "",
    note: "",
  });

  React.useEffect(() => {
    setForm((current) => ({ ...current, spaceId: spaces[0]?.id ?? current.spaceId }));
  }, [spaces]);

  return (
    <div className="page">
      <PageHeader
        eyebrow="Shopping coordination"
        title="Auto-generated replenishment plus manual household asks."
        description="Forecast-driven entries appear automatically from low stock and expiry risk. Manual additions can come from any role, including contributors."
      />

      {!permissions.canPurchaseShopping ? (
        <PermissionNotice>
          Contributor mode can add needs to the list, but only owners and members can mark items purchased and push stock back into inventory.
        </PermissionNotice>
      ) : null}

      <div className="twoColumn">
        <Surface strong>
          <SectionTitle title="Current shopping list" action={<span className="badge badgeOrange">{shoppingList.length} entries</span>} />
          {shoppingList.length ? (
            <div className="stack">
              {shoppingList.map((entry) => {
                const matchingItem = entry.inventoryItemId
                  ? state.inventoryItems.find((item) => item.id === entry.inventoryItemId)
                  : undefined;
                const space = entry.spaceId ? spaces.find((item) => item.id === entry.spaceId) : undefined;
                return (
                  <div key={entry.id} className="listItem">
                    <div className="sectionHeading">
                      <div className="stackTight">
                        <strong>{entry.productName}</strong>
                        <span className="muted">
                          {categoryLabels[entry.category]} • {formatQuantity(entry.quantity, entry.quantityUnit)}
                        </span>
                      </div>
                      <span className={`badge ${entry.source === "manual" ? "badgeNeutral" : "badgeOrange"}`}>{entry.source}</span>
                    </div>
                    <p className="muted">{entry.note ?? "No note provided."}</p>
                    <p className="muted">
                      {space ? `${space.name}` : "Flexible destination"} {matchingItem ? "• matched to tracked inventory" : "• purchase will create or restock inventory"}
                    </p>
                    <button className="button" type="button" onClick={() => actions.purchaseShoppingItem(entry)} disabled={!permissions.canPurchaseShopping}>
                      Mark purchased
                    </button>
                  </div>
                );
              })}
            </div>
          ) : (
            <EmptyState title="Shopping list is clear" description="No low-stock or manual entries are waiting right now." />
          )}
        </Surface>

        <Surface>
          <SectionTitle title="Add manual request" action={<span className="badge badgeGreen">Contributor-safe</span>} />
          <form
            className="stack"
            onSubmit={(event) => {
              event.preventDefault();
              actions.addShoppingItem({ ...form, quantity: Number(form.quantity) });
              setForm((current) => ({ ...current, productName: "", note: "" }));
            }}
          >
            <div className="formGrid">
              <label className="field">
                <span className="label">Product</span>
                <input className="input" value={form.productName} onChange={(event) => setForm((current) => ({ ...current, productName: event.target.value }))} required />
              </label>
              <label className="field">
                <span className="label">Category</span>
                <select className="select" value={form.category} onChange={(event) => setForm((current) => ({ ...current, category: event.target.value as typeof current.category }))}>
                  {Object.entries(categoryLabels).map(([value, label]) => (
                    <option key={value} value={value}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span className="label">Quantity</span>
                <input className="input" type="number" min="1" value={form.quantity} onChange={(event) => setForm((current) => ({ ...current, quantity: Number(event.target.value) }))} />
              </label>
              <label className="field">
                <span className="label">Unit</span>
                <select className="select" value={form.quantityUnit} onChange={(event) => setForm((current) => ({ ...current, quantityUnit: event.target.value as typeof current.quantityUnit }))}>
                  <option value="count">count</option>
                  <option value="ml">ml</option>
                  <option value="g">g</option>
                  <option value="percentage">percentage</option>
                </select>
              </label>
              <label className="field">
                <span className="label">Space</span>
                <select className="select" value={form.spaceId} onChange={(event) => setForm((current) => ({ ...current, spaceId: event.target.value }))}>
                  {spaces.map((space) => (
                    <option key={space.id} value={space.id}>
                      {space.name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <label className="field">
              <span className="label">Note</span>
              <textarea className="textarea" value={form.note} onChange={(event) => setForm((current) => ({ ...current, note: event.target.value }))} placeholder="Why it matters, brand preference, or who asked for it." />
            </label>
            <button className="buttonSecondary" type="submit">
              Add to list
            </button>
          </form>
        </Surface>
      </div>
    </div>
  );
}
