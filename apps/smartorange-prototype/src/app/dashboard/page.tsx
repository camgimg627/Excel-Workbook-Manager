"use client";

import Link from "next/link";
import { formatDateTime } from "@/lib/date";
import { categoryLabels, formatQuantity } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";
import { MetricCard, PageHeader, ProgressMeter, SectionTitle, Surface } from "@/components/ui";

export default function DashboardPage() {
  const { activeHousehold, activeUser, forecasts, metrics, notifications, recentActivity, spaceSnapshots, state } =
    useDemoSession();
  const expiringItems = forecasts
    .filter((entry) => state.inventoryItems.find((item) => item.id === entry.inventoryItemId)?.expiryDate)
    .slice(0, 3);

  return (
    <div className="page">
      <PageHeader
        eyebrow="Household command center"
        title={`${activeHousehold.name} at a glance`}
        description={`Built for ${activeUser.displayName} to understand what is getting low, what expires soon, and where the next shopping or recipe action should happen.`}
        actions={
          <>
            <Link className="button" href="/shopping-list">
              Open shopping list
            </Link>
            <Link className="buttonGhost" href="/capture">
              Capture new stock
            </Link>
          </>
        }
      />

      <div className="metricGrid">
        <MetricCard label="Visible items" value={`${metrics.visibleItemCount}`} detail="Current inventory in this role view." />
        <MetricCard label="Low stock" value={`${metrics.lowStockCount}`} detail="Items already in the replenishment zone." />
        <MetricCard label="Expiring soon" value={`${metrics.expiringCount}`} detail="Needs use-it-now attention." />
        <MetricCard label="Shared goods" value={`${metrics.sharedCount}`} detail="Shared or household-visible items in circulation." />
        <MetricCard label="Avg coverage" value={`${metrics.averageDays}d`} detail="Average days remaining across visible stock." />
      </div>

      <div className="twoColumn">
        <Surface strong>
          <SectionTitle title="Forecast timeline" action={<span className="badge badgeOrange">Depletion curves</span>} />
          <div className="stack">
            {forecasts.slice(0, 5).map((forecast) => {
              const item = state.inventoryItems.find((entry) => entry.id === forecast.inventoryItemId);
              return (
                <div key={forecast.inventoryItemId} className="listItem">
                  <div className="sectionHeading">
                    <div className="stackTight">
                      <strong>{forecast.productName}</strong>
                      <span className="muted">
                        {categoryLabels[forecast.category]} • {formatQuantity(forecast.quantityValue, forecast.quantityUnit)}
                      </span>
                    </div>
                    <span className={`badge ${forecast.lowStock ? "badgeOrange" : "badgeGreen"}`}>
                      {forecast.daysRemaining}d left
                    </span>
                  </div>
                  <ProgressMeter
                    value={(forecast.quantityValue / Math.max(item?.quantityFullValue ?? forecast.quantityFullValue, 1)) * 100}
                    label={`${forecast.productName} stock level`}
                    detail={`Confidence: ${forecast.confidence} • Replenish by ${new Date(forecast.replenishmentDate).toLocaleDateString()}`}
                  />
                </div>
              );
            })}
          </div>
        </Surface>

        <div className="stack">
          <Surface>
            <SectionTitle title="Notifications inbox" />
            <ul className="list">
              {notifications.map((entry) => (
                <li key={entry.id} className="listItem">
                  <div className="sectionHeading">
                    <strong>{entry.title}</strong>
                    <span className={`badge ${entry.severity === "urgent" ? "badgeOrange" : "badgeNeutral"}`}>
                      {entry.severity}
                    </span>
                  </div>
                  <p className="muted">{entry.body}</p>
                </li>
              ))}
            </ul>
          </Surface>

          <Surface>
            <SectionTitle title="Quick actions" />
            <div className="sectionGrid">
              <Link className="listItem" href="/inventory">
                <strong>Audit inventory</strong>
                <p className="muted">Search, filter, and update quantities or visibility.</p>
              </Link>
              <Link className="listItem" href="/recipes">
                <strong>Cook from what you already have</strong>
                <p className="muted">Use recipe logs to deduct ingredients and refresh forecasts.</p>
              </Link>
              <Link className="listItem" href="/assistant">
                <strong>Ask the assistant</strong>
                <p className="muted">Grounded answers over the same seeded data layer.</p>
              </Link>
            </div>
          </Surface>
        </div>
      </div>

      <div className="twoColumn">
        <Surface>
          <SectionTitle title="Spaces snapshot" />
          <div className="sectionGrid">
            {spaceSnapshots.map((entry) => (
              <div key={entry.space.id} className="listItem">
                <strong>{entry.space.name}</strong>
                <p className="muted">
                  {entry.itemCount} tracked items • {entry.lowStockCount} low-stock calls
                </p>
              </div>
            ))}
          </div>
        </Surface>

        <Surface>
          <SectionTitle title="Recent activity" />
          <ul className="list">
            {recentActivity.map((entry) => (
              <li key={entry.id} className="listItem">
                <strong>{entry.label}</strong>
                <p className="muted">{entry.detail}</p>
                <span className="muted mono">{formatDateTime(entry.at)}</span>
              </li>
            ))}
          </ul>
        </Surface>
      </div>

      <Surface>
        <SectionTitle title="Expiry watch" action={<span className="badge badgeNeutral">Use before waste</span>} />
        <div className="sectionGrid">
          {expiringItems.length ? (
            expiringItems.map((forecast) => {
              const item = state.inventoryItems.find((entry) => entry.id === forecast.inventoryItemId);
              const expiryDate = item?.expiryDate;
              return (
                <div key={forecast.inventoryItemId} className="listItem">
                  <strong>{forecast.productName}</strong>
                  <p className="muted">
                    {expiryDate ? `Expires ${new Date(expiryDate).toLocaleDateString()}` : "No expiry date recorded."}
                  </p>
                  <p className="muted">Suggested next move: cook or use it before the next replenishment cycle.</p>
                </div>
              );
            })
          ) : (
            <div className="emptyState">No immediate expiry risk in the current role view.</div>
          )}
        </div>
      </Surface>
    </div>
  );
}
