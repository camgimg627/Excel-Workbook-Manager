export function addDays(dateLike: string | Date, days: number): Date {
  const date = new Date(dateLike);
  date.setDate(date.getDate() + days);
  return date;
}

export function daysUntil(dateLike: string | Date): number {
  const now = new Date();
  const target = new Date(dateLike);
  return Math.ceil((target.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
}

export function formatDate(dateLike: string | Date): string {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
  }).format(new Date(dateLike));
}

export function formatDateTime(dateLike: string | Date): string {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(dateLike));
}

export function formatRelativeDays(days: number): string {
  if (days <= 0) {
    return "today";
  }
  if (days === 1) {
    return "in 1 day";
  }
  return `in ${days} days`;
}

export function todayIso(): string {
  return new Date().toISOString();
}
