"use client";

import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import IconoCapex from "./iconos/IconoCapex";
import IconoOpex from "./iconos/IconoOpex";

interface Enlace {
  href: string;
  etiqueta: string;
}

interface Categoria {
  clave: "capex" | "opex";
  etiqueta: string;
  icono: (props: { size?: number }) => React.ReactElement;
  enlaces: Enlace[];
}

/** Las dos categorías principales del dashboard — misma estructura de páginas que ya
 *  existía en Sidebar, solo horizontal (arriba) en vez de vertical (a la izquierda), para
 *  liberar ancho para el contenido. Ningún href cambia. */
const CATEGORIAS: Categoria[] = [
  {
    clave: "capex",
    etiqueta: "CAPEX",
    icono: IconoCapex,
    enlaces: [
      { href: "/", etiqueta: "Dashboard" },
      { href: "/bd-capex", etiqueta: "Detalle BD_CAPEX" },
      { href: "/facturas", etiqueta: "Facturas" },
    ],
  },
  {
    clave: "opex",
    etiqueta: "OPEX",
    icono: IconoOpex,
    enlaces: [
      { href: "/opex", etiqueta: "Dashboard" },
      { href: "/opex/presupuesto", etiqueta: "Presupuesto" },
      { href: "/opex/facturas", etiqueta: "Facturas" },
    ],
  },
];

interface Props {
  /** Sin contraseña configurada (dev local) — no muestra el botón de salir. */
  local: boolean;
}

export default function TopNav({ local }: Props) {
  const pathname = usePathname();
  const router = useRouter();
  const categoriaActiva = pathname.startsWith("/opex") ? "opex" : "capex";
  const catActivaObj = CATEGORIAS.find((c) => c.clave === categoriaActiva)!;

  async function cerrarSesion() {
    await fetch("/api/login", { method: "DELETE" });
    router.push("/login");
    router.refresh();
  }

  return (
    <div style={{ borderBottom: "1px solid var(--borde)", background: "var(--card)" }}>
      <div className="max-w-[1800px] px-6 flex items-center gap-4 flex-wrap">
        <span className="text-xs font-bold tracking-wide shrink-0" style={{ color: "var(--texto-suave)" }}>
          CONTROL DE GASTOS
        </span>

        <nav className="flex items-center gap-1">
          {CATEGORIAS.map((cat) => {
            const activa = categoriaActiva === cat.clave;
            const Icono = cat.icono;
            return (
              <Link
                key={cat.clave}
                href={cat.enlaces[0].href}
                className="flex items-center gap-1.5 text-sm"
                style={{
                  padding: "0.7rem 0.75rem",
                  fontWeight: activa ? 700 : 500,
                  color: activa ? "var(--acento-fuerte)" : "var(--texto-suave)",
                  borderBottom: activa ? "3px solid var(--acento)" : "3px solid transparent",
                }}
              >
                <Icono size={16} />
                {cat.etiqueta}
              </Link>
            );
          })}
        </nav>

        <div className="w-px self-stretch my-2" style={{ background: "var(--borde)" }} />

        <nav className="flex items-center gap-1">
          {catActivaObj.enlaces.map((e) => {
            const enlaceActivo = pathname === e.href;
            return (
              <Link
                key={e.href}
                href={e.href}
                className="text-xs rounded-md"
                style={{
                  padding: "0.4rem 0.65rem",
                  fontWeight: enlaceActivo ? 600 : 400,
                  color: enlaceActivo ? "var(--acento-fuerte)" : "var(--texto-suave)",
                  background: enlaceActivo ? "var(--acento-suave)" : "transparent",
                }}
              >
                {e.etiqueta}
              </Link>
            );
          })}
        </nav>

        {!local && (
          <button
            onClick={cerrarSesion}
            className="text-xs font-medium hover:underline ml-auto shrink-0"
            style={{ color: "var(--acento)" }}
          >
            Cerrar sesión
          </button>
        )}
      </div>
    </div>
  );
}
