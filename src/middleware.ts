import { NextRequest, NextResponse } from "next/server";
import {
  APP_PASSWORD_CONFIGURADA,
  APP_PASSWORD_LECTURA_CONFIGURADA,
  COOKIE_ACCESO,
  COOKIE_ACCESO_LECTURA,
  tokenEsperado,
  tokenEsperadoLectura,
} from "@/lib/appAuth";

/**
 * Rutas SIEMPRE públicas, para cualquiera (con o sin ninguna cookie) — a propósito:
 * cualquiera con el link puede VER y DESCARGAR facturas sin pedir nada, mientras que
 * registrar o editar una factura (rutas de escritura, más abajo) exige acceso completo.
 * Ojo: son coincidencias EXACTAS (o de ese subpath puntual) a propósito — así
 * "/api/opex/facturas" (GET, lista) queda libre pero "/api/opex/facturas/registrar"
 * (POST, escribe) NO, aunque comparta el mismo prefijo.
 */
const RUTAS_SIEMPRE_LIBRES = [
  /^\/facturas\/?$/,
  /^\/opex\/facturas\/?$/,
  /^\/api\/facturas\/?$/,
  /^\/api\/facturas\/exportar\/?$/,
  /^\/api\/opex\/facturas\/?$/,
  /^\/api\/opex\/facturas\/exportar\/?$/,
  /^\/api\/auth\/nivel\/?$/, // para que cualquier página sepa su propio nivel de acceso
];

/**
 * Rutas de solo VISTA que además puede ver quien entró con la contraseña de solo
 * lectura (nunca de escritura) — Dashboard y Presupuesto OPEX, tal como se pidió: esa
 * persona puede mirar, nunca manipular. BD_CAPEX y todo lo demás quedan fuera a propósito.
 */
const RUTAS_SOLO_LECTURA = [
  /^\/$/,
  /^\/opex\/?$/,
  /^\/opex\/presupuesto\/?$/,
  /^\/api\/capex\/?$/,
  /^\/api\/opex\/?$/,
];

function coincideAlguna(patrones: RegExp[], pathname: string): boolean {
  return patrones.some((patron) => patron.test(pathname));
}

/** /api/opex/mes-cierre es GET (solo lee) o POST (cierra un mes de verdad) en la misma
 *  ruta — el nivel de solo lectura solo puede usar el GET. */
function permitidoParaLectura(pathname: string, metodo: string): boolean {
  if (coincideAlguna(RUTAS_SOLO_LECTURA, pathname)) return true;
  if (pathname === "/api/opex/mes-cierre" && metodo === "GET") return true;
  return false;
}

// Sin ninguna contraseña configurada (dev local sin la nube), no hay nada que proteger:
// se deja pasar todo, igual que antes de agregar esto.
export default async function middleware(req: NextRequest) {
  if (!APP_PASSWORD_CONFIGURADA) return NextResponse.next();

  const { pathname } = req.nextUrl;
  if (coincideAlguna(RUTAS_SIEMPRE_LIBRES, pathname)) return NextResponse.next();

  const cookieCompleto = req.cookies.get(COOKIE_ACCESO)?.value;
  if (cookieCompleto && cookieCompleto === (await tokenEsperado())) return NextResponse.next();

  if (APP_PASSWORD_LECTURA_CONFIGURADA) {
    const cookieLectura = req.cookies.get(COOKIE_ACCESO_LECTURA)?.value;
    if (cookieLectura && cookieLectura === (await tokenEsperadoLectura())) {
      if (permitidoParaLectura(pathname, req.method)) return NextResponse.next();
      // Ya inició sesión (con la contraseña de solo lectura), solo que esta ruta no es
      // para ese nivel — a un módulo que sí puede ver, no al login (no le falta iniciar
      // sesión, le falta permiso).
      const url = req.nextUrl.clone();
      url.pathname = "/";
      url.search = "";
      return NextResponse.redirect(url);
    }
  }

  const url = req.nextUrl.clone();
  url.pathname = "/login";
  url.searchParams.set("volver", pathname);
  return NextResponse.redirect(url);
}

// Protege todas las páginas y rutas /api/* excepto login y assets estáticos.
export const config = {
  matcher: ["/((?!api/login|login|_next/static|_next/image|favicon.ico).*)"],
};
