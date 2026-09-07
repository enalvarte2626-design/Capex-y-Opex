import { NextResponse } from "next/server";
import { cookies } from "next/headers";
import {
  APP_PASSWORD_CONFIGURADA,
  APP_PASSWORD_LECTURA_CONFIGURADA,
  COOKIE_ACCESO,
  COOKIE_ACCESO_LECTURA,
  tokenEsperado,
  tokenEsperadoLectura,
} from "@/lib/appAuth";

export const dynamic = "force-dynamic";

/**
 * Le dice a una página con qué nivel de acceso está entrando quien la mira — para que,
 * por ejemplo, Facturas oculte el formulario de registro y los campos editables a quien
 * no tenga acceso completo (sin contraseña configurada, todo el mundo es "completo": la
 * app sigue igual que antes de agregar esto).
 */
export async function GET() {
  if (!APP_PASSWORD_CONFIGURADA) {
    return NextResponse.json({ nivel: "completo" });
  }

  const jar = await cookies();
  const cookieCompleto = jar.get(COOKIE_ACCESO)?.value;
  if (cookieCompleto && cookieCompleto === (await tokenEsperado())) {
    return NextResponse.json({ nivel: "completo" });
  }

  if (APP_PASSWORD_LECTURA_CONFIGURADA) {
    const cookieLectura = jar.get(COOKIE_ACCESO_LECTURA)?.value;
    if (cookieLectura && cookieLectura === (await tokenEsperadoLectura())) {
      return NextResponse.json({ nivel: "lectura" });
    }
  }

  return NextResponse.json({ nivel: "ninguno" });
}
