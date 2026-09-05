import { NextRequest, NextResponse } from "next/server";
import { APP_PASSWORD_CONFIGURADA, COOKIE_ACCESO, tokenEsperado } from "@/lib/appAuth";

// Sin APP_PASSWORD configurada (dev local sin la nube), no hay nada que proteger: se deja
// pasar todo, igual que antes de agregar esto.
export default async function middleware(req: NextRequest) {
  if (!APP_PASSWORD_CONFIGURADA) return NextResponse.next();

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
