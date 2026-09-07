/**
 * Acceso al dashboard con contraseñas compartidas (sin cuentas individuales) —
 * alternativa al login con Microsoft Entra ID cuando no hay permisos de administración
 * en Azure para agregar el Redirect URI que ese flujo necesita.
 *
 * Hay DOS niveles de acceso, cada uno con su propia contraseña y su propia cookie:
 * - "completo" (APP_PASSWORD): puede ver y editar todo.
 * - "lectura" (APP_PASSWORD_LECTURA, opcional): solo puede VER el Dashboard y
 *   Presupuesto OPEX — nunca editar nada. Pensado para compartir el link con alguien
 *   externo sin darle permiso de tocar los datos. Si esta variable no está configurada,
 *   ese nivel simplemente no existe (nadie puede entrar con él).
 *
 * El login solo se activa si hay al menos una contraseña configurada. En desarrollo
 * local, sin ninguna variable, la app sigue funcionando exactamente igual que antes de
 * agregar esto: sin pantalla de acceso.
 *
 * Ninguna cookie guarda la contraseña en texto plano: cada una guarda un hash SHA-256
 * de "contraseña + AUTH_SECRET" (el mismo AUTH_SECRET ya generado para esto). Se compara
 * con 'crypto.subtle', disponible tanto en el runtime de Node como en el de Edge (el
 * middleware corre en Edge), así no hace falta forzar runtime de Node ahí.
 */

export const COOKIE_ACCESO = "capex_acceso";
export const COOKIE_ACCESO_LECTURA = "capex_acceso_lectura";
export const APP_PASSWORD_CONFIGURADA = Boolean(process.env.APP_PASSWORD);
export const APP_PASSWORD_LECTURA_CONFIGURADA = Boolean(process.env.APP_PASSWORD_LECTURA);

async function sha256Hex(texto: string): Promise<string> {
  const datos = new TextEncoder().encode(texto);
  const hash = await crypto.subtle.digest("SHA-256", datos);
  return Array.from(new Uint8Array(hash))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

/** Valor que debe tener la cookie de sesión (acceso completo) para considerarse válida. */
export async function tokenEsperado(): Promise<string> {
  const secreto = process.env.AUTH_SECRET ?? "";
  return sha256Hex(`${process.env.APP_PASSWORD ?? ""}:${secreto}`);
}

/** Igual que `tokenEsperado`, pero para la cookie de solo lectura. */
export async function tokenEsperadoLectura(): Promise<string> {
  const secreto = process.env.AUTH_SECRET ?? "";
  return sha256Hex(`${process.env.APP_PASSWORD_LECTURA ?? ""}:${secreto}`);
}

/** Compara la contraseña ingresada contra la de acceso completo (server-only). */
export function contrasenaValida(intento: string): boolean {
  return Boolean(process.env.APP_PASSWORD) && intento === process.env.APP_PASSWORD;
}

/** Compara la contraseña ingresada contra la de solo lectura (server-only). */
export function contrasenaValidaLectura(intento: string): boolean {
  return Boolean(process.env.APP_PASSWORD_LECTURA) && intento === process.env.APP_PASSWORD_LECTURA;
}
