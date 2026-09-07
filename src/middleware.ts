import { NextRequest, NextResponse } from "next/server";
import { APP_PASSWORD_CONFIGURADA, COOKIE_ACCESO, tokenEsperado } from "@/lib/appAuth";

/**
 * Módulo de Facturas (CAPEX y OPEX) — queda libre, sin pedir contraseña, a propósito:
 * cualquiera con el link puede entrar directo a registrar/ver facturas, mientras que el
 * resto de la app (Dashboards, BD_CAPEX, Presupuesto) sigue detrás de la contraseña.
 */
const RUTAS_LIBRES = [
  /^\/facturas(\/|$)/,
  /^\/opex\/facturas(\/|$)/,
  /^\/api\/facturas(\/|$)/,
  /^\/api\/opex\/facturas(\/|$)/,
  /^\/api\/opex\/agregar-linea$/, // misma acción de registro que /api/opex/facturas/registrar
];

function esRutaLibre(pathname: string): boolean {
  return RUTAS_LIBRES.some((patron) => patron.test(pathname));
}

// Sin APP_PASSWORD configurada (dev local sin la nube), no hay nada que proteger: se deja
// pasar todo, igual que antes de agregar esto.
export default async function middleware(req: NextRequest) {
  if (!APP_PASSWORD_CONFIGURADA) return NextResponse.next();
  if (esRutaLibre(req.nextUrl.pathname)) return NextResponse.next();

  const cookie = req.cookies.get(COOKIE_ACCESO)?.value;
  if (cookie && cookie === (await tokenEsperado())) return NextResponse.next();

  const url = req.nextUrl.clone();
  url.pathname = "/login";
  url.searchParams.set("volver", req.nextUrl.pathname);
  return NextResponse.redirect(url);
}

// Protege todas las páginas y rutas /api/* excepto login y assets estáticos.
export const config = {
  matcher: ["/((?!api/login|login|_next/static|_next/image|favicon.ico).*)"],
};
