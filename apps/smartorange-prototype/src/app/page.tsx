"use client";

import Link from "next/link";
import { PageHeader, Surface } from "@/components/ui";
import { categoryLabels } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";

const routeHighlights = [
  {
    title: "Dashboard",
    body: "Forecasted runout, expiry risk, space-level visibility, and a household inbox of prompts.",
  },
  {
    title: "Capture",
    body: "Manual logging plus mocked barcode, receipt, and browser-extension flows in one operational hub.",
  },
  {
    title: "Inventory",
    body: "Searchable stock views with quantity modes, visibility rules, and quick use/update controls.",
  },
  {
    title: "Recipes + Assistant",
    body: "Recipe-driven deduction and a grounded conversational layer over the same structured demo state.",
  },
];

export default function HomePage() {
  const { state, activeHousehold, actions } = useDemoSession();

  return (
    <div className="page">
      <section className="hero">
        <div className="heroGrid">
          <div className="stack">
            <span className="eyebrow">V1 foundation prototype</span>
            <h2 className="displayTitle">A household data layer that finally knows what is left, not just what exists.</h2>
            <p className="lede">
              SmartOrange turns pantry shelves, bathroom cabinets, and shared staples into structured, household-aware
              inventory. This prototype focuses on the V1 loop: capture, visibility, depletion, replenishment, recipes,
              and quick AI-style answers.
            </p>
            <div className="cluster">
              <Link className="button" href="/dashboard">
                Open dashboard
              </Link>
              <Link className="buttonSecondary" href="/capture">
                Try data capture
              </Link>
            </div>
          </div>

          <Surface strong>
            <p className="label">What this preview proves</p>
            <div className="stack">
              <div>
                <strong>Depletion-aware inventory</strong>
                <p className="muted">Units and percentage-based items forecast differently, but live together in one system.</p>
              </div>
              <div>
                <strong>Household-aware permissions</strong>
                <p className="muted">Switch roles instantly to compare owner, member, and contributor visibility.</p>
              </div>
              <div>
                <strong>Actionable, not ornamental</strong>
                <p className="muted">Every screen is wired to local demo state so a purchase, cook log, or usage entry changes the rest of the app.</p>
              </div>
            </div>
          </Surface>
        </div>
      </section>

      <PageHeader
        eyebrow="Demo households"
        title="Choose the household story you want to prototype against."
        description="Each seeded household mirrors a V1 persona with different visibility, spaces, and replenishment pressure."
      />

      <div className="sectionGrid">
        {state.households.map((household) => (
          <Surface key={household.id} strong={household.id === activeHousehold.id}>
            <span className="badge badgeNeutral">{household.type}</span>
            <h3 className="sectionTitle">{household.name}</h3>
            <p className="muted">{household.summary}</p>
            <div className="cluster">
              <button type="button" className="buttonGhost" onClick={() => actions.switchHousehold(household.id)}>
                Focus this household
              </button>
              <Link className="buttonSecondary" href="/dashboard" onClick={() => actions.switchHousehold(household.id)}>
                Preview flow
              </Link>
            </div>
          </Surface>
        ))}
      </div>

      <div className="twoColumn">
        <Surface>
          <div className="sectionHeading">
            <h3 className="sectionTitle">Prototype coverage</h3>
            <span className="badge badgeOrange">Routes wired</span>
          </div>
          <div className="sectionGrid">
            {routeHighlights.map((entry) => (
              <div key={entry.title} className="listItem">
                <strong>{entry.title}</strong>
                <p className="muted">{entry.body}</p>
              </div>
            ))}
          </div>
        </Surface>

        <Surface>
          <div className="sectionHeading">
            <h3 className="sectionTitle">V1 category scope</h3>
            <span className="badge badgeGreen">Locked to roadmap</span>
          </div>
          <div className="stack">
            {Object.entries(categoryLabels).map(([key, label]) => (
              <div key={key} className="listItem">
                <strong>{label}</strong>
                <p className="muted">
                  {key === "food"
                    ? "Pantry staples, produce, dairy, recipe ingredients, and recurring meal coverage."
                    : key === "household_supplies"
                      ? "Cleaning, paper goods, laundry, and shared utility stock."
                      : "Shampoo, health items, bathroom supplies, and private care inventory."}
                </p>
              </div>
            ))}
          </div>
        </Surface>
      </div>
    </div>
  );
}
