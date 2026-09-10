/**
 * Agrupa nombres de proveedor que en realidad son el mismo, para sugerirlos sin
 * duplicados en los selectores de Facturas (CAPEX y OPEX).
 *
 * Dos variantes de mayúsculas/espacios ya se consideraban iguales ("Go daddy" = "Go
 * Daddy" = "GO DADDY "), pero un caso real quedaba afuera: "Metrica" vs "Metrica Sac"
 * — el mismo proveedor, solo que a veces se escribe con la razón social completa. Para
 * cubrir eso, la clave de agrupación además quita el sufijo de tipo de empresa al final
 * del nombre (SAC, SA, SRL, EIRL, LTDA, INC, etc.) — nunca en medio del nombre, así que
 * "SA" de "USA Travel" o similar no se toca por accidente.
 */
const SUFIJOS_EMPRESA = new Set([
  "SAC",
  "SA",
  "SAA",
  "SRL",
  "SCRL",
  "EIRL",
  "SAJ",
  "LTDA",
  "LTD",
  "LLC",
  "INC",
  "CORP",
  "CO",
]);

/** Exportada para que otros lugares (ej. la sugerencia de RUC por proveedor) agrupen con
 *  el mismo criterio exacto que `agruparProveedores`. */
export function claveNormalizada(nombre: string): string {
  const palabras = nombre.toLowerCase().replace(/\s+/g, " ").trim().split(" ");
  // Quita como mucho un sufijo de tipo de empresa al final ("Metrica Sac" -> "metrica"),
  // nunca deja el nombre vacío.
  if (palabras.length > 1) {
    const ultima = palabras[palabras.length - 1].replace(/\./g, "").toUpperCase();
    if (SUFIJOS_EMPRESA.has(ultima)) palabras.pop();
  }
  return palabras.join(" ");
}

/**
 * A partir de una lista de nombres tal como se escribieron (con repetidos), agrupa por
 * `claveNormalizada` y devuelve, de cada grupo, la forma exacta más usada — ordenado
 * alfabéticamente. Es lo que se muestra en el <datalist> del campo Proveedor.
 */
export function agruparProveedores(nombres: (string | undefined | null)[]): string[] {
  const conteoPorClave = new Map<string, Map<string, number>>();
  for (const nombreCrudo of nombres) {
    const nombre = nombreCrudo?.trim();
    if (!nombre) continue;
    const clave = claveNormalizada(nombre);
    const formas = conteoPorClave.get(clave) ?? new Map<string, number>();
    formas.set(nombre, (formas.get(nombre) ?? 0) + 1);
    conteoPorClave.set(clave, formas);
  }
  const resultado = Array.from(conteoPorClave.values()).map((formas) => {
    let mejor = "";
    let mejorConteo = -1;
    for (const [forma, veces] of formas) {
      if (veces > mejorConteo) {
        mejor = forma;
        mejorConteo = veces;
      }
    }
    return mejor;
  });
  return resultado.sort((a, b) => a.localeCompare(b, "es"));
}

/**
 * A partir de las facturas ya registradas, arma un mapa (proveedor normalizado -> RUC
 * más usado con ese proveedor) — para sugerir el RUC en automático apenas se elige un
 * proveedor ya conocido. Un proveedor nuevo (o uno que nunca se registró con RUC) no
 * aparece en el mapa, así que el campo queda vacío para completarlo a mano.
 */
export function mapaRucPorProveedor(entradas: { proveedor: string | undefined | null; ruc: string | undefined | null }[]): Map<string, string> {
  const conteoPorClave = new Map<string, Map<string, number>>();
  for (const { proveedor, ruc } of entradas) {
    const nombre = proveedor?.trim();
    const rucLimpio = ruc?.trim();
    if (!nombre || !rucLimpio) continue;
    const clave = claveNormalizada(nombre);
    const formas = conteoPorClave.get(clave) ?? new Map<string, number>();
    formas.set(rucLimpio, (formas.get(rucLimpio) ?? 0) + 1);
    conteoPorClave.set(clave, formas);
  }
  const resultado = new Map<string, string>();
  for (const [clave, formas] of conteoPorClave) {
    let mejor = "";
    let mejorConteo = -1;
    for (const [ruc, veces] of formas) {
      if (veces > mejorConteo) {
        mejor = ruc;
        mejorConteo = veces;
      }
    }
    resultado.set(clave, mejor);
  }
  return resultado;
}
