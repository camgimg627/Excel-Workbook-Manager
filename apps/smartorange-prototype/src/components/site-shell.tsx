"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import styles from "@/components/site-shell.module.css";
import { useDemoSession } from "@/lib/session-context";

const navItems = [
  { href: "/", label: "Onboarding" },
  { href: "/dashboard", label: "Dashboard" },
  { href: "/capture", label: "Capture" },
  { href: "/inventory", label: "Inventory" },
  { href: "/shopping-list", label: "Shopping List" },
  { href: "/recipes", label: "Recipes" },
  { href: "/assistant", label: "Assistant" },
];

export function SiteShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const { state, activeHousehold, activeMember, actions } = useDemoSession();
  const householdMembers = state.householdMembers.filter((entry) => entry.householdId === activeHousehold.id);

  return (
    <div className={styles.frame}>
      <div className={styles.topBar}>
        <div className={styles.brandRow}>
          <div>
            <div className="eyebrow">SmartOrange Prototype</div>
            <h1 className={styles.brandTitle}>Household inventory, finally made conversational.</h1>
            <p className={styles.brandSub}>{activeHousehold.summary}</p>
          </div>
          <div className={styles.chipRow}>
            <span className="badge badgeOrange">{activeHousehold.type}</span>
            <span className="badge badgeGreen">{activeMember.role}</span>
            <button type="button" className="buttonGhost" onClick={() => actions.reset()}>
              Reset demo
            </button>
          </div>
        </div>

        <div className={styles.controlRow}>
          <div className={styles.selectors}>
            <label className={styles.selectWrap}>
              <span className={styles.selectLabel}>Household</span>
              <select
                className={styles.select}
                value={activeHousehold.id}
                onChange={(event) => actions.switchHousehold(event.target.value)}
              >
                {state.households.map((entry) => (
                  <option key={entry.id} value={entry.id}>
                    {entry.name}
                  </option>
                ))}
              </select>
            </label>

            <label className={styles.selectWrap}>
              <span className={styles.selectLabel}>Role View</span>
              <select
                className={styles.select}
                value={activeMember.id}
                onChange={(event) => actions.switchMember(event.target.value)}
              >
                {householdMembers.map((entry) => {
                  const user = state.users.find((userEntry) => userEntry.id === entry.userId);
                  return (
                    <option key={entry.id} value={entry.id}>
                      {user?.displayName} ({entry.role})
                    </option>
                  );
                })}
              </select>
            </label>
          </div>

          <div className={styles.chipRow}>
            <span className="badge badgeNeutral mono">localhost:3112</span>
            <span className="badge badgeNeutral">V1 categories only</span>
          </div>
        </div>

        <nav className={styles.navRow} aria-label="Primary">
          {navItems.map((item) => {
            const active = pathname === item.href;
            return (
              <Link
                key={item.href}
                href={item.href}
                className={`${styles.navLink} ${active ? styles.navLinkActive : ""}`}
              >
                {item.label}
              </Link>
            );
          })}
        </nav>
      </div>

      <main className={styles.content}>{children}</main>
    </div>
  );
}
