"use client";

import * as React from "react";
import { Category, QuantityUnit, Visibility } from "@/lib/domain";
import { categoryLabels } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";
import { EmptyState, PageHeader, PermissionNotice, SectionTitle, Surface } from "@/components/ui";

export default function CapturePage() {
  const { activeHousehold, permissions, spaces, state, actions } = useDemoSession();
  const householdDrafts = state.captureDrafts.filter((entry) => entry.householdId === activeHousehold.id);
  const [manualForm, setManualForm] = React.useState<{
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
    retailer: string;
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
    retailer: "Manual log",
  });

  React.useEffect(() => {
    setManualForm((current) => ({
      ...current,
      spaceId: spaces[0]?.id ?? "",
    }));
  }, [spaces]);

  const submitManualForm = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    actions.addInventory({
      ...manualForm,
      quantityValue: Number(manualForm.quantityValue),
      quantityFullValue: Number(manualForm.quantityFullValue),
      expiryDate: manualForm.expiryDate || undefined,
    });
    setManualForm((current) => ({
      ...current,
      name: "",
      notes: "",
      expiryDate: "",
    }));
  };

  return (
    <div className="page">
      <PageHeader
        eyebrow="Data capture"
        title="Four intake paths, one structured household model."
        description="Manual entry is fully interactive. Barcode, receipt, and browser-extension flows are simulated with realistic drafts that resolve into inventory updates and purchase events."
      />

      {!permissions.canCaptureToInventory ? (
        <PermissionNotice>
          Contributor mode can preview capture flows, but only owners and members can turn them into inventory. Contributors can still add follow-up needs on the shopping list.
        </PermissionNotice>
      ) : null}

      <div className="twoColumn">
        <Surface strong>
          <SectionTitle title="Manual log" action={<span className="badge badgeGreen">Interactive</span>} />
          <form className="stack" onSubmit={submitManualForm}>
            <div className="formGrid">
              <label className="field">
                <span className="label">Item name</span>
                <input
                  className="input"
                  value={manualForm.name}
                  onChange={(event) => setManualForm((current) => ({ ...current, name: event.target.value }))}
                  placeholder="Dishwasher pods"
                  required
                />
              </label>
              <label className="field">
                <span className="label">Category</span>
                <select
                  className="select"
                  value={manualForm.category}
                  onChange={(event) =>
                    setManualForm((current) => ({
                      ...current,
                      category: event.target.value as typeof current.category,
                    }))
                  }
                >
                  {Object.entries(categoryLabels).map(([value, label]) => (
                    <option key={value} value={value}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span className="label">Subcategory</span>
                <input
                  className="input"
                  value={manualForm.subcategory}
                  onChange={(event) => setManualForm((current) => ({ ...current, subcategory: event.target.value }))}
                />
              </label>
              <label className="field">
                <span className="label">Space</span>
                <select
                  className="select"
                  value={manualForm.spaceId}
                  onChange={(event) => setManualForm((current) => ({ ...current, spaceId: event.target.value }))}
                >
                  {spaces.map((space) => (
                    <option key={space.id} value={space.id}>
                      {space.name}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span className="label">Quantity</span>
                <input
                  className="input"
                  type="number"
                  min="0"
                  step="1"
                  value={manualForm.quantityValue}
                  onChange={(event) => setManualForm((current) => ({ ...current, quantityValue: Number(event.target.value) }))}
                />
              </label>
              <label className="field">
                <span className="label">Unit</span>
                <select
                  className="select"
                  value={manualForm.quantityUnit}
                  onChange={(event) =>
                    setManualForm((current) => ({
                      ...current,
                      quantityUnit: event.target.value as typeof current.quantityUnit,
                    }))
                  }
                >
                  <option value="count">count</option>
                  <option value="ml">ml</option>
                  <option value="g">g</option>
                  <option value="percentage">percentage</option>
                </select>
              </label>
              <label className="field">
                <span className="label">Full amount</span>
                <input
                  className="input"
                  type="number"
                  min="1"
                  step="1"
                  value={manualForm.quantityFullValue}
                  onChange={(event) =>
                    setManualForm((current) => ({ ...current, quantityFullValue: Number(event.target.value) }))
                  }
                />
              </label>
              <label className="field">
                <span className="label">Visibility</span>
                <select
                  className="select"
                  value={manualForm.visibility}
                  onChange={(event) =>
                    setManualForm((current) => ({
                      ...current,
                      visibility: event.target.value as typeof current.visibility,
                    }))
                  }
                >
                  <option value="private">Private</option>
                  <option value="shared">Shared</option>
                  <option value="household">Household</option>
                </select>
              </label>
              <label className="field">
                <span className="label">Expiry date</span>
                <input
                  className="input"
                  type="date"
                  value={manualForm.expiryDate}
                  onChange={(event) => setManualForm((current) => ({ ...current, expiryDate: event.target.value }))}
                />
              </label>
            </div>

            <label className="field">
              <span className="label">Notes</span>
              <textarea
                className="textarea"
                value={manualForm.notes}
                onChange={(event) => setManualForm((current) => ({ ...current, notes: event.target.value }))}
                placeholder="Optional context: bulk buy, allergy-safe, opened already..."
              />
            </label>

            <button className="button" type="submit" disabled={!permissions.canCaptureToInventory}>
              Save to inventory
            </button>
          </form>
        </Surface>

        <Surface>
          <SectionTitle title="Mocked capture channels" action={<span className="badge badgeOrange">Try samples</span>} />
          <div className="stack">
            <div className="listItem">
              <strong>Barcode scan</strong>
              <p className="muted">Single product autofill with catalog confidence.</p>
              <button className="buttonGhost" type="button" onClick={() => actions.queueCapture("barcode")}>
                Queue barcode sample
              </button>
            </div>
            <div className="listItem">
              <strong>Receipt / photo parse</strong>
              <p className="muted">Transforms one image into multiple inventory candidates with confidence scores.</p>
              <button className="buttonGhost" type="button" onClick={() => actions.queueCapture("receipt")}>
                Queue receipt sample
              </button>
            </div>
            <div className="listItem">
              <strong>Browser extension capture</strong>
              <p className="muted">Passive retailer purchase capture before the delivery even lands on the porch.</p>
              <button className="buttonGhost" type="button" onClick={() => actions.queueCapture("browser")}>
                Queue browser sample
              </button>
            </div>
          </div>
        </Surface>
      </div>

      <Surface>
        <SectionTitle title="Pending capture drafts" action={<span className="badge badgeNeutral">{householdDrafts.length} active</span>} />
        {householdDrafts.length ? (
          <div className="sectionGrid">
            {householdDrafts.map((draft) => (
              <div key={draft.id} className="listItem">
                <div className="sectionHeading">
                  <div className="stackTight">
                    <strong>{draft.title}</strong>
                    <span className="muted">
                      {draft.retailer} • {draft.method.replace("_", " ")}
                    </span>
                  </div>
                  <span className="badge badgeOrange">{draft.items.length} items</span>
                </div>
                <p className="muted">{draft.note}</p>
                <ul className="list">
                  {draft.items.map((item) => {
                    const space = spaces.find((entry) => entry.id === item.spaceId);
                    return (
                      <li key={item.id} className="listItem">
                        <strong>{item.productName}</strong>
                        <p className="muted">
                          {item.quantity} {item.quantityUnit} • {space?.name ?? "Unknown space"} • {Math.round(item.confidence * 100)}%
                          confidence
                        </p>
                      </li>
                    );
                  })}
                </ul>
                <div className="cluster">
                  <button className="button" type="button" onClick={() => actions.acceptCapture(draft.id)} disabled={!permissions.canCaptureToInventory}>
                    Accept into inventory
                  </button>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <EmptyState
            title="No active drafts"
            description="Queue a barcode, receipt, or browser sample above to simulate passive intake."
          />
        )}
      </Surface>
    </div>
  );
}
