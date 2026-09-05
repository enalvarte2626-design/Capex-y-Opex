/**
 * Acceso al dashboard con una sola contraseña compartida (sin cuentas individuales) —
 * alternativa al login con Microsoft Entra ID cuando no hay permisos de administración
 * en Azure para agregar el Redirect URI que ese flujo necesita.
 *
 * El login solo se activa si hay una contraseña configurada ('APP_PASSWORD'). En
 * desarrollo local, sin esa variable, la app sigue funcionando exactamente igual que
 * antes de agregar esto: sin pantalla de acceso.
 *
 * La cookie de sesión nunca guarda la contraseña en texto plano: guarda un hash
 * SHA-256 de "contraseña + AUTH_SECRET" (el mismo AUTH_SECRET ya generado para esto).
 * Se compara con 'crypto.subtle', disponible tanto en el runtime de Node como en el de
 * Edge (el middleware corre en Edge), asi no hace falta forzar runtime de Node ahi.
 */

export const COOKIE_ACCESO = "capex_acceso";
export const APP_PASSWORD_CONFIGURADA = Boolean(process.env.APP_PASSWORD);

async function sha256Hex(texto: string): Promise<string> {
  const datos = new TextEncoder().encode(texto);
  const hash = await crypto.subtle.digest("SHA-256", datos);
  return Array.from(new Uint8Array(hash))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

/** Valor que debe tener la cookie de sesión para considerarse válida. */
export async function tokenEsperado(): Promise<string> {
  const secreto = process.env.AUTH_SECRET ?? "";
  return sha256Hex(`${process.env.APP_PASSWORD ?? ""}:${secreto}`);
}

/** Compara la contraseña ingresada contra la configurada (comparación simple, server-only). */
export function contrasenaValida(intento: string): boolean {
  return Boolean(process.env.APP_PASSWORD) && intento === process.env.APP_PASSWORD;
}
