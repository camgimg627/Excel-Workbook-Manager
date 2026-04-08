import type { Metadata } from "next";
import { Fraunces, IBM_Plex_Mono, Space_Grotesk } from "next/font/google";
import { SiteShell } from "@/components/site-shell";
import { DemoSessionProvider } from "@/lib/session-context";
import "./globals.css";

const display = Fraunces({
  subsets: ["latin"],
  variable: "--font-display",
});

const body = Space_Grotesk({
  subsets: ["latin"],
  variable: "--font-body",
});

const mono = IBM_Plex_Mono({
  subsets: ["latin"],
  variable: "--font-mono",
  weight: ["400", "500"],
});

export const metadata: Metadata = {
  title: "SmartOrange Prototype",
  description: "Household inventory prototype focused on visibility, depletion forecasting, and shopping coordination.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body className={`${display.variable} ${body.variable} ${mono.variable}`}>
        <DemoSessionProvider>
          <SiteShell>{children}</SiteShell>
        </DemoSessionProvider>
      </body>
    </html>
  );
}
