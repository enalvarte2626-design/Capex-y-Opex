import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import { crearHojaSiNoExiste, escribirCelda, leerCelda } from "./sharepoint";

/**
 * "Año de presupuesto activo" guardado EN EL PROPIO EXCEL (hoja "Config App", misma hoja
 * donde ya vive el mes de cierre — ver mesCierreConfig.ts — pero en la fila 2 en vez de
 * la 1, para no pisarla) — no en una variable de entorno ni en código, porque el botón
 * "Aprobar y activar" de Planificación tiene que poder cambiarlo él mismo, sin que nadie
 * tenga que entrar a Azure ni desplegar nada cada año que empieza.
 *
 * CAPEX y OPEX viven en archivos de Excel distintos, así que cada uno guarda su propio
 * año activo en su propia hoja "Config App" — normalmente van a coincidir, pero nada
 * obliga a aprobar los dos el mismo día.
 */
const HOJA_CONFIG = "Config App";
const CELDA_ETIQUETA = "A2";
const CELDA_VALOR = "B2";

/** Año activo actual — si la hoja "Config App" todavía no tiene nada en B2 (nunca se
 *  aprobó ningún año desde el botón), usa `porDefecto`. */
export async function leerAnioActivo(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  porDefecto: number
): Promise<number> {
  try {
    const valor = await leerCelda(config, archivo, HOJA_CONFIG, CELDA_VALOR);
    if (Number.isInteger(valor) && valor >= 2000 && valor <= 2200) return valor;
  } catch {
    // La hoja "Config App" todavía no existe.
  }
  return porDefecto;
}

/** Guarda un nuevo año activo — se llama solo desde "Aprobar y activar", nunca a mano. */
export async function escribirAnioActivo(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  anio: number
): Promise<void> {
  await crearHojaSiNoExiste(config, archivo, HOJA_CONFIG);
  await escribirCelda(
    config,
    archivo,
    HOJA_CONFIG,
    CELDA_ETIQUETA,
    'Año de presupuesto activo — se cambia solo al presionar "Aprobar y activar" en Planificación, no a mano acá.'
  );
  await escribirCelda(config, archivo, HOJA_CONFIG, CELDA_VALOR, anio);
}
