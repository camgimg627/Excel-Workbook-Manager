"use client";

import * as React from "react";
import { Category, QuantityUnit, Visibility } from "@/lib/domain";
import { categoryLabels, formatQuantity, visibilityLabels } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";
import { EmptyState, PageHeader, PermissionNotice, ProgressMeter, SectionTitle, Surface } from "@/components/ui";

export default function InventoryPage() {
  const { forecasts, permissions, spaces, state, visibleInventory, actions } = useDemoSession();
  const [query, setQuery] = React.useState("");
  const [filters, setFilters] = React.useState({
    category: "all",
    spaceId: "all",
    visibility: "all",
    quantityMode: "all",
    lowStockOnly: false,
  });
  const [selectedItemId, setSelectedItemId] = React.useState<string | null>(visibleInventory[0]?.id ?? null);
  const [showCreate, setShowCreate] = React.useState(false);
  const [createForm, setCreateForm] = React.useState<{
    name: string;
    category: Category;
    subcategory: string;
    quantityValue: number;
    quantityUnit: QuantityUnit;
    quantityFullValue: number;
    spaceId: string;
    visibility: Visibility;
    expiryDate: string;
    notes: string;
  }>({
    name: "",
    category: "food",
    subcategory: "pantry",
    quantityValue: 1,
    quantityUnit: "count",
    quantityFullValue: 1,
    spaceId: spaces[0]?.id ?? "",
    visibility: "shared",
    expiryDate: "",
    notes: "",
  });
  const deferredQuery = React.useDeferredValue(query);

  React.useEffect(() => {
    setCreateForm((current) => ({ ...current, spaceId: spaces[0]?.id ?? current.spaceId }));
  }, [spaces]);

  const filteredItems = visibleInventory.filter((item) => {
    const product = state.products.find((entry) => entry.id === item.productId);
    const matchesQuery = `${product?.name ?? ""} ${product?.subcategory ?? ""} ${item.notes ?? ""}`
      .toLowerCase()
      .includes(deferredQuery.toLowerCase());
    const matchesCategory = filters.category === "all" || product?.category === filters.category;
    const matchesSpace = filters.spaceId === "all" || item.spaceId === filters.spaceId;
    const matchesVisibility = filters.visibility === "all" || item.visibility === filters.visibility;
    const matchesMode =
      filters.quantityMode === "all" ||
      (filters.quantityMode === "percentage" ? item.quantityUnit === "percentage" : item.quantityUnit !== "percentage");
    const lowStock = forecasts.some((entry) => entry.inventoryItemId === item.id && entry.lowStock);

    return matchesQuery && matchesCategory && matchesSpace && matchesVisibility && matchesMode && (!filters.lowStockOnly || lowStock);
  });

  React.useEffect(() => {
    if (!filteredItems.length) {
      setSelectedItemId(null);
      return;
    }

    if (!filteredItems.some((entry) => entry.id === selectedItemId)) {
      setSelectedItemId(filteredItems[0].id);
    }
  }, [filteredItems, selectedItemId]);

  const selectedItem =
    filteredItems.find((entry) => entry.id === selectedItemId) ??
    visibleInventory.find((entry) => entry.id === selectedItemId) ??
    null;
  const selectedProduct = selectedItem ? state.products.find((entry) => entry.id === selectedItem.productId) : null;
  const selectedSpace = selectedItem ? spaces.find((entry) => entry.id === selectedItem.spaceId) : null;
  const selectedForecast = selectedItem ? forecasts.find((entry) => entry.inventoryItemId === selectedItem.id) : null;

  const [editForm, setEditForm] = React.useState<{
    spaceId: string;
    visibility: Visibility;
    quantityValue: number;
    quantityFullValue: number;
    expiryDate: string;
    notes: string;
  }>({
    spaceId: "",
    visibility: "shared",
    quantityValue: 0,
    quantityFullValue: 0,
    expiryDate: "",
    notes: "",
  });
  const [useAmount, setUseAmount] = React.useState(1);

  React.useEffect(() => {
    if (!selectedItem) {
      return;
    }

    setEditForm({
      spaceId: selectedItem.spaceId,
      visibility: selectedItem.visibility,
      quantityValue: selectedItem.quantityValue,
      quantityFullValue: selectedItem.quantityFullValue,
      expiryDate: selectedItem.expiryDate ?? "",
      notes: selectedItem.notes ?? "",
    });
    setUseAmount(selectedItem.quantityUnit === "percentage" ? 5 : 1);
  }, [selectedItem]);

  return (
    <div className="page">
      <PageHeader
        eyebrow="Inventory manager"
        title="Search, filter, and act on the current household state."
        description="This view stays aligned to the role switcher: owners can see everything, members see private items they own, and contributors get a limited read-only shared view."
        actions={
          permissions.canEditInventory ? (
            <button
              type="button"
              className="button"
              onClick={() => {
                setShowCreate(true);
                setSelectedItemId(null);
              }}
            >
              Add inventory item
            </button>
          ) : undefined
        }
      />

      {!permissions.canEditInventory ? (
        <PermissionNotice>
          Contributor mode can review shared inventory and use it to build the shopping list, but editing quantities or visibility is disabled.
        </PermissionNotice>
      ) : null}

      <div className="twoColumn">
        <Surface strong>
          <SectionTitle title="Inventory view" action={<span className="badge badgeNeutral">{filteredItems.length} results</span>} />
          <div className="stack">
            <div className="formGrid">
              <label className="field">
                <span className="label">Search</span>
                <input className="input" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search by name, notes, or subtype" />
              </label>
              <label className="field">
                <span className="label">Category</span>
                <select className="select" value={filters.category} onChange={(event) => setFilters((current) => ({ ...current, category: event.target.value }))}>
                  <option value="all">All categories</option>
                  {Object.entries(categoryLabels).map(([value, label]) => (
                    <option key={value} value={value}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span className="label">Space</span>
                <select className="select" value={filters.spaceId} onChange={(event) => setFilters((current) => ({ ...current, spaceId: event.target.value }))}>
                  <option value="all">All spaces</option>
                  {spaces.map((space) => (
                    <option key={space.id} value={space.id}>
                      {space.name}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span className="label">Visibility</span>
                <select className="select" value={filters.visibility} onChange={(event) => setFilters((current) => ({ ...current, visibility: event.target.value }))}>
                  <option value="all">Any visibility</option>
                  <option value="private">Private</option>
                  <option value="shared">Shared</option>
                  <option value="household">Household</option>
                </select>
              </label>
              <label className="field">
                <span className="label">Quantity mode</span>
                <select className="select" value={filters.quantityMode} onChange={(event) => setFilters((current) => ({ ...current, quantityMode: event.target.value }))}>
                  <option value="all">All</option>
                  <option value="unit">Units / measurable</option>
                  <option value="percentage">Percentage-based</option>
                </select>
              </label>
              <label className="field">
                <span className="label">Low stock only</span>
                <select
                  className="select"
                  value={filters.lowStockOnly ? "yes" : "no"}
                  onChange={(event) => setFilters((current) => ({ ...current, lowStockOnly: event.target.value === "yes" }))}
                >
                  <option value="no">Show all</option>
                  <option value="yes">Only flagged items</option>
                </select>
              </label>
            </div>

            {filteredItems.length ? (
              <div className="tableWrap">
                <table className="table">
                  <thead>
                    <tr>
                      <th>Item</th>
                      <th>Category</th>
                      <th>Space</th>
                      <th>Visibility</th>
                      <th>Quantity</th>
                      <th>Forecast</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredItems.map((item) => {
                      const product = state.products.find((entry) => entry.id === item.productId);
                      const space = spaces.find((entry) => entry.id === item.spaceId);
                      const forecast = forecasts.find((entry) => entry.inventoryItemId === item.id);
                      return (
                        <tr key={item.id} onClick={() => { setSelectedItemId(item.id); setShowCreate(false); }}>
                          <td>
                            <strong>{product?.name ?? "Unknown item"}</strong>
                            <div className="muted">{item.notes ?? "No notes recorded"}</div>
                          </td>
                          <td>{product ? categoryLabels[product.category] : "Unknown"}</td>
                          <td>{space?.name ?? "Unknown"}</td>
                          <td>{visibilityLabels[item.visibility]}</td>
                          <td>{formatQuantity(item.quantityValue, item.quantityUnit)}</td>
                          <td>{forecast ? `${forecast.daysRemaining}d` : "No signal"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ) : (
              <EmptyState title="No inventory matches these filters." description="Broaden the filters or switch households to explore another demo state." />
            )}
          </div>
        </Surface>

        <Surface>
          {showCreate ? (
            <div className="drawer">
              <SectionTitle title="Create new inventory item" />
              <form
                className="stack"
                onSubmit={(event) => {
                  event.preventDefault();
                  actions.addInventory({
                    ...createForm,
                    quantityValue: Number(createForm.quantityValue),
                    quantityFullValue: Number(createForm.quantityFullValue),
                    expiryDate: createForm.expiryDate || undefined,
                  });
                  setShowCreate(false);
                  setCreateForm((current) => ({ ...current, name: "", notes: "", expiryDate: "" }));
                }}
              >
                <div className="formGrid">
                  <label className="field">
                    <span className="label">Name</span>
                    <input className="input" value={createForm.name} onChange={(event) => setCreateForm((current) => ({ ...current, name: event.target.value }))} required />
                  </label>
                  <label className="field">
                    <span className="label">Category</span>
                    <select className="select" value={createForm.category} onChange={(event) => setCreateForm((current) => ({ ...current, category: event.target.value as typeof current.category }))}>
                      {Object.entries(categoryLabels).map(([value, label]) => (
                        <option key={value} value={value}>
                          {label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field">
                    <span className="label">Subtype</span>
                    <input className="input" value={createForm.subcategory} onChange={(event) => setCreateForm((current) => ({ ...current, subcategory: event.target.value }))} />
                  </label>
                  <label className="field">
                    <span className="label">Quantity</span>
                    <input className="input" type="number" min="0" value={createForm.quantityValue} onChange={(event) => setCreateForm((current) => ({ ...current, quantityValue: Number(event.target.value) }))} />
                  </label>
                  <label className="field">
                    <span className="label">Unit</span>
                    <select className="select" value={createForm.quantityUnit} onChange={(event) => setCreateForm((current) => ({ ...current, quantityUnit: event.target.value as typeof current.quantityUnit }))}>
                      <option value="count">count</option>
                      <option value="ml">ml</option>
                      <option value="g">g</option>
                      <option value="percentage">percentage</option>
                    </select>
                  </label>
                  <label className="field">
                    <span className="label">Full amount</span>
                    <input className="input" type="number" min="1" value={createForm.quantityFullValue} onChange={(event) => setCreateForm((current) => ({ ...current, quantityFullValue: Number(event.target.value) }))} />
                  </label>
                  <label className="field">
                    <span className="label">Space</span>
                    <select className="select" value={createForm.spaceId} onChange={(event) => setCreateForm((current) => ({ ...current, spaceId: event.target.value }))}>
                      {spaces.map((space) => (
                        <option key={space.id} value={space.id}>
                          {space.name}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field">
                    <span className="label">Visibility</span>
                    <select className="select" value={createForm.visibility} onChange={(event) => setCreateForm((current) => ({ ...current, visibility: event.target.value as typeof current.visibility }))}>
                      <option value="private">Private</option>
                      <option value="shared">Shared</option>
                      <option value="household">Household</option>
                    </select>
                  </label>
                  <label className="field">
                    <span className="label">Expiry</span>
                    <input className="input" type="date" value={createForm.expiryDate} onChange={(event) => setCreateForm((current) => ({ ...current, expiryDate: event.target.value }))} />
                  </label>
                </div>
                <label className="field">
                  <span className="label">Notes</span>
                  <textarea className="textarea" value={createForm.notes} onChange={(event) => setCreateForm((current) => ({ ...current, notes: event.target.value }))} />
                </label>
                <div className="cluster">
                  <button className="button" type="submit" disabled={!permissions.canEditInventory}>
                    Save item
                  </button>
                  <button className="buttonGhost" type="button" onClick={() => setShowCreate(false)}>
                    Cancel
                  </button>
                </div>
              </form>
            </div>
          ) : selectedItem && selectedProduct ? (
            <div className="drawer">
              <SectionTitle title={selectedProduct.name} action={<span className="badge badgeNeutral">{visibilityLabels[selectedItem.visibility]}</span>} />
              <p className="muted">
                {categoryLabels[selectedProduct.category]} • {selectedSpace?.name ?? "Unknown space"}
              </p>
              <ProgressMeter
                value={(selectedItem.quantityValue / Math.max(selectedItem.quantityFullValue, 1)) * 100}
                label={`Current stock: ${formatQuantity(selectedItem.quantityValue, selectedItem.quantityUnit)}`}
                detail={selectedForecast ? `${selectedForecast.daysRemaining} days remaining • ${selectedForecast.confidence} signal` : "No forecast yet"}
              />
              <div className="stack">
                <label className="field">
                  <span className="label">Space</span>
                  <select className="select" value={editForm.spaceId} onChange={(event) => setEditForm((current) => ({ ...current, spaceId: event.target.value }))}>
                    {spaces.map((space) => (
                      <option key={space.id} value={space.id}>
                        {space.name}
                      </option>
                    ))}
                  </select>
                </label>
                <div className="formGrid">
                  <label className="field">
                    <span className="label">Visibility</span>
                    <select className="select" value={editForm.visibility} onChange={(event) => setEditForm((current) => ({ ...current, visibility: event.target.value as typeof current.visibility }))}>
                      <option value="private">Private</option>
                      <option value="shared">Shared</option>
                      <option value="household">Household</option>
                    </select>
                  </label>
                  <label className="field">
                    <span className="label">Quantity</span>
                    <input className="input" type="number" min="0" value={editForm.quantityValue} onChange={(event) => setEditForm((current) => ({ ...current, quantityValue: Number(event.target.value) }))} />
                  </label>
                  <label className="field">
                    <span className="label">Full amount</span>
                    <input className="input" type="number" min="1" value={editForm.quantityFullValue} onChange={(event) => setEditForm((current) => ({ ...current, quantityFullValue: Number(event.target.value) }))} />
                  </label>
                  <label className="field">
                    <span className="label">Expiry</span>
                    <input className="input" type="date" value={editForm.expiryDate} onChange={(event) => setEditForm((current) => ({ ...current, expiryDate: event.target.value }))} />
                  </label>
                </div>
                <label className="field">
                  <span className="label">Notes</span>
                  <textarea className="textarea" value={editForm.notes} onChange={(event) => setEditForm((current) => ({ ...current, notes: event.target.value }))} />
                </label>
                <div className="cluster">
                  <button className="button" type="button" onClick={() => actions.updateInventory(selectedItem.id, { ...editForm, expiryDate: editForm.expiryDate || undefined })} disabled={!permissions.canEditInventory}>
                    Save updates
                  </button>
                </div>
              </div>
              <Surface>
                <SectionTitle title="Log usage" />
                <div className="cluster">
                  <input className="input" style={{ maxWidth: "120px" }} type="number" min="0" value={useAmount} onChange={(event) => setUseAmount(Number(event.target.value))} />
                  <button className="buttonSecondary" type="button" onClick={() => actions.useInventory(selectedItem.id, Number(useAmount))} disabled={!permissions.canEditInventory}>
                    Use some
                  </button>
                </div>
              </Surface>
            </div>
          ) : (
            <EmptyState title="Pick an item to inspect" description="Select an inventory row to open its details, forecast, and quick actions." />
          )}
        </Surface>
      </div>
    </div>
  );
}
