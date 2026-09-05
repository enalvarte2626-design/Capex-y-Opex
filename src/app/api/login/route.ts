import { NextResponse } from "next/server";
import { COOKIE_ACCESO, contrasenaValida, tokenEsperado } from "@/lib/appAuth";

export const dynamic = "force-dynamic";

/** Valida la contraseña y, si es correcta, deja la cookie de sesión (hash, nunca en claro). */
export async function POST(req: Request) {
  const body = await req.json().catch(() => ({}));
  const password = typeof body?.password === "string" ? body.password : "";

  if (!contrasenaValida(password)) {
    return NextResponse.json({ error: "Contraseña incorrecta." }, { status: 401 });
  }

  const res = NextResponse.json({ ok: true });
  res.cookies.set(COOKIE_ACCESO, await tokenEsperado(), {
    httpOnly: true,
    secure: true,
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 30, // 30 días
  });
  return res;
}

/** Cierra sesión: borra la cookie de acceso. */
export async function DELETE() {
  const res = NextResponse.json({ ok: true });
  res.cookies.delete(COOKIE_ACCESO);
  return res;
}
