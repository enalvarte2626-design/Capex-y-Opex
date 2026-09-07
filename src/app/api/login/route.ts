import { NextResponse } from "next/server";
import {
  COOKIE_ACCESO,
  COOKIE_ACCESO_LECTURA,
  contrasenaValida,
  contrasenaValidaLectura,
  tokenEsperado,
  tokenEsperadoLectura,
} from "@/lib/appAuth";

export const dynamic = "force-dynamic";

/** Valida la contraseña (completa o de solo lectura) y deja la cookie que corresponda —
 *  nunca las dos a la vez, para que quede claro con qué nivel entró. */
export async function POST(req: Request) {
  const body = await req.json().catch(() => ({}));
  const password = typeof body?.password === "string" ? body.password : "";

  if (contrasenaValida(password)) {
    const res = NextResponse.json({ ok: true, nivel: "completo" });
    res.cookies.set(COOKIE_ACCESO, await tokenEsperado(), {
      httpOnly: true,
      secure: true,
      sameSite: "lax",
      path: "/",
      maxAge: 60 * 60 * 24 * 30, // 30 días
    });
    res.cookies.delete(COOKIE_ACCESO_LECTURA);
    return res;
  }

  if (contrasenaValidaLectura(password)) {
    const res = NextResponse.json({ ok: true, nivel: "lectura" });
    res.cookies.set(COOKIE_ACCESO_LECTURA, await tokenEsperadoLectura(), {
      httpOnly: true,
      secure: true,
      sameSite: "lax",
      path: "/",
      maxAge: 60 * 60 * 24 * 30, // 30 días
    });
    res.cookies.delete(COOKIE_ACCESO);
    return res;
  }

  return NextResponse.json({ error: "Contraseña incorrecta." }, { status: 401 });
}

/** Cierra sesión: borra ambas cookies de acceso. */
export async function DELETE() {
  const res = NextResponse.json({ ok: true });
  res.cookies.delete(COOKIE_ACCESO);
  res.cookies.delete(COOKIE_ACCESO_LECTURA);
  return res;
}
