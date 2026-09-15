import type { Metadata } from "next";
import "./globals.css";
import AppShell from "@/components/AppShell";
import { AnioProvider } from "@/components/AnioProvider";
import { APP_PASSWORD_CONFIGURADA } from "@/lib/appAuth";
import { anioPresupuestoActual } from "@/lib/anio";

// Sin esto, Next podía pre-renderizar este layout como HTML estático en el build (varias
// páginas no tienen datos propios del servidor, solo cargan del Excel del lado del
// cliente) — y entonces `anioPresupuestoActual()` quedaría "congelado" con el valor que
// tuviera ANIO_PRESUPUESTO en el momento de compilar, sin importar qué diga la variable
// de entorno después en Azure. "force-dynamic" obliga a que el layout (y por lo tanto el
// año que reparte a toda la app) se calcule de nuevo en cada visita — mismo criterio que
// ya usan las rutas de API, que nunca cachean nada del Excel.
export const dynamic = "force-dynamic";

// Función (no un objeto fijo) para que lea la variable de entorno ANIO_PRESUPUESTO en
// cada arranque del servidor, no solo una vez al compilar — así cambiarla en Azure no
// necesita un despliegue nuevo, solo reiniciar la app (que Azure ya hace solo).
export async function generateMetadata(): Promise<Metadata> {
  return {
    title: "Control Presupuestal TI",
    description: `Control de gasto e inversión CAPEX/OPEX ${anioPresupuestoActual()} — Expertia`,
  };
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es" className="h-full">
      <body className="min-h-full">
        <AnioProvider anio={anioPresupuestoActual()}>
          <AppShell local={!APP_PASSWORD_CONFIGURADA}>{children}</AppShell>
        </AnioProvider>
      </body>
    </html>
  );
}
