import type { Metadata } from "next";
import "./globals.css";
import AppShell from "@/components/AppShell";
import { APP_PASSWORD_CONFIGURADA } from "@/lib/appAuth";

export const metadata: Metadata = {
  title: "Dashboard CAPEX",
  description: "Control de gasto e inversión CAPEX 2026 — Expertia",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es" className="h-full">
      <body className="min-h-full">
        <AppShell local={!APP_PASSWORD_CONFIGURADA}>{children}</AppShell>
      </body>
    </html>
  );
}
