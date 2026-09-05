"use client";

import { usePathname } from "next/navigation";
import TopNav from "./TopNav";

interface Props {
  local: boolean;
  children: React.ReactNode;
}

/** Envoltorio de toda la app: barra de navegación arriba (categorías CAPEX/OPEX,
 *  horizontal — no ocupa ancho lateral) + área principal con un título que se adapta a
 *  la sección activa. Ninguna página cambia — siguen recibiendo exactamente el mismo
 *  `children` que recibían antes. */
export default function AppShell({ local, children }: Props) {
  const pathname = usePathname();
  const esOpex = pathname.startsWith("/opex");
  const esLogin = pathname === "/login";

  if (esLogin) {
    // La pantalla de login ya tiene su propio layout centrado de pantalla completa.
    return <>{children}</>;
  }

  return (
    <div className="flex flex-col min-h-full">
      <TopNav local={local} />
      <header className="border-b" style={{ borderColor: "var(--borde)", background: "var(--card)" }}>
        <div className="max-w-[1800px] px-6 py-4">
          <h1 className="text-xl font-semibold">
            {esOpex ? "PRESUPUESTO OPEX TI - 2026" : "PRESUPUESTO CAPEX TI - 2026"}
          </h1>
        </div>
      </header>
      <main className="max-w-[1800px] w-full px-6 py-8">{children}</main>
    </div>
  );
}
