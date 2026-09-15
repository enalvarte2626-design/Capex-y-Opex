import type { Metadata } from "next";
import "./globals.css";
import AppShell from "@/components/AppShell";
import { AnioProvider } from "@/components/AnioProvider";
import { APP_PASSWORD_CONFIGURADA } from "@/lib/appAuth";
import { anioActivoCapex, anioActivoOpex } from "@/lib/anio";

// Sin esto, Next podía pre-renderizar este layout como HTML estático en el build (varias
// páginas no tienen datos propios del servidor, solo cargan del Excel del lado del
// cliente) — y entonces el año de presupuesto quedaría "congelado" con lo que hubiera en
// el Excel al momento de compilar, sin reflejar nunca un "Aprobar y activar" posterior.
// "force-dynamic" obliga a que el layout (y por lo tanto el año que reparte a toda la
// app) se calcule de nuevo en cada visita — mismo criterio que ya usan las rutas de API,
// que nunca cachean nada del Excel.
export const dynamic = "force-dynamic";

export const metadata: Metadata = {
  title: "Control Presupuestal TI",
  description: "Control de gasto e inversión CAPEX/OPEX — Expertia",
};

export default async function RootLayout({ children }: { children: React.ReactNode }) {
  // Cada módulo vive en su propio archivo de Excel, así que cada uno tiene su propio año
  // activo (leído de la hoja "Config App" de cada archivo) — ver lib/anio.ts.
  const [anioCapex, anioOpex] = await Promise.all([anioActivoCapex(), anioActivoOpex()]);

  return (
    <html lang="es" className="h-full">
      <body className="min-h-full">
        <AnioProvider capex={String(anioCapex)} opex={String(anioOpex)}>
          <AppShell local={!APP_PASSWORD_CONFIGURADA}>{children}</AppShell>
        </AnioProvider>
      </body>
    </html>
  );
}
