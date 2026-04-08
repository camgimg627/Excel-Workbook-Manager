import styles from "@/components/ui.module.css";

function join(...values: Array<string | false | null | undefined>) {
  return values.filter(Boolean).join(" ");
}

export function Surface({
  children,
  strong = false,
  className,
}: {
  children: React.ReactNode;
  strong?: boolean;
  className?: string;
}) {
  return <section className={join("surface", strong && "surfaceStrong", styles.surface, className)}>{children}</section>;
}

export function PageHeader({
  eyebrow,
  title,
  description,
  actions,
}: {
  eyebrow: string;
  title: string;
  description: string;
  actions?: React.ReactNode;
}) {
  return (
    <div className="pageHeader">
      <div className="stackTight">
        <span className="eyebrow">{eyebrow}</span>
        <h2 className="pageTitle">{title}</h2>
        <p className="lede">{description}</p>
      </div>
      {actions ? <div className="cluster">{actions}</div> : null}
    </div>
  );
}

export function MetricCard({
  label,
  value,
  detail,
}: {
  label: string;
  value: string;
  detail: string;
}) {
  return (
    <Surface className={styles.metricCard}>
      <p className="label">{label}</p>
      <div className="metricValue">{value}</div>
      <p className={styles.description}>{detail}</p>
    </Surface>
  );
}

export function SectionTitle({
  title,
  action,
}: {
  title: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="sectionHeading">
      <h3 className="sectionTitle">{title}</h3>
      {action}
    </div>
  );
}

export function ProgressMeter({
  value,
  label,
  detail,
}: {
  value: number;
  label: string;
  detail?: string;
}) {
  return (
    <div className="stackTight">
      <div className={styles.progressMeta}>
        <strong>{label}</strong>
        <span>{Math.round(Math.min(Math.max(value, 0), 100))}%</span>
      </div>
      <div className="progressTrack">
        <div className="progressFill" style={{ width: `${Math.min(Math.max(value, 0), 100)}%` }} />
      </div>
      {detail ? <span className="muted">{detail}</span> : null}
    </div>
  );
}

export function EmptyState({
  title,
  description,
  action,
}: {
  title: string;
  description: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="emptyState stackTight">
      <strong>{title}</strong>
      <span>{description}</span>
      {action}
    </div>
  );
}

export function PermissionNotice({ children }: { children: React.ReactNode }) {
  return <div className="notice">{children}</div>;
}
