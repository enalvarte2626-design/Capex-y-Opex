import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import { ErrorSharePoint, crearArchivo, crearHojaSiNoExiste, descargarContenido, escribirCelda, leerCelda, listarNombresCarpeta } from "./sharepoint";
import { MES_CIERRE_POR_DEFECTO } from "./opex-constantes";

/**
 * "Mes de cierre" (1-12) guardado EN EL PROPIO EXCEL — no en localStorage ni en código —
 * para que sea un solo valor real, compartido por cualquiera que use la app: el mismo
 * que separa Real de Forecast en el Dashboard, y el que decide si una factura nueva
 * suma o no al Gasto Real al registrarla (ver /api/opex/facturas/registrar).
 *
 * Antes esto era una constante fija en el código (MES_CIERRE_POR_DEFECTO) que solo yo
 * podía cambiar editando y desplegando. Ahora vive en una hoja chica del Excel
 * ("Config App"), así que el usuario lo cierra él mismo desde un botón en Presupuesto
 * OPEX, sin depender de un cambio de código para cada mes que se cierre.
 */
const HOJA_CONFIG = "Config App";
const CELDA_ETIQUETA = "A1";
const CELDA_VALOR = "B1";

/** Mes de cierre actual — si la hoja "Config App" todavía no existe (nadie cerró ningún
 *  mes desde el botón todavía), usa `valorPorDefecto` como punto de partida (OPEX pasa
 *  MES_CIERRE_POR_DEFECTO; CAPEX pasa lo que sugiera el nombre del archivo actual, ver
 *  /api/capex/mes-cierre, para no "perder" de golpe los meses que ya estaban cerrados
 *  según el patrón "N+M" del archivo). */
export async function leerMesCierre(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  valorPorDefecto: number = MES_CIERRE_POR_DEFECTO
): Promise<number> {
  try {
    const valor = await leerCelda(config, archivo, HOJA_CONFIG, CELDA_VALOR);
    if (Number.isInteger(valor) && valor >= 1 && valor <= 12) return valor;
  } catch {
    // La hoja "Config App" todavía no existe — nadie cerró ningún mes desde el botón.
  }
  return valorPorDefecto;
}

/** "Control Capex Forecast 8+4.xlsm" → 8 — el patrón "N+M" del nombre del archivo en
 *  vivo de CAPEX, usado como valor de arranque de `leerMesCierre` la primera vez (antes
 *  de que exista la hoja "Config App" en ese archivo). Un archivo sin ese patrón se
 *  trata como 0 meses cerrados. */
export function mesesCerradosPorNombreArchivo(nombre: string): number {
  const m = nombre.match(/(\d+)\s*\+\s*(\d+)/);
  return m ? Number(m[1]) : 0;
}

/** Guarda un nuevo mes de cierre — crea la hoja "Config App" la primera vez que hace falta. */
export async function escribirMesCierre(config: ConfiguracionSharePoint, archivo: ArchivoResuelto, mes: number): Promise<void> {
  await crearHojaSiNoExiste(config, archivo, HOJA_CONFIG);
  await escribirCelda(
    config,
    archivo,
    HOJA_CONFIG,
    CELDA_ETIQUETA,
    "Mes de cierre (1-12) — el último mes con Gasto Real ya cerrado. Se edita solo desde el botón \"Cerrar mes\" en Presupuesto OPEX, no a mano acá."
  );
  await escribirCelda(config, archivo, HOJA_CONFIG, CELDA_VALOR, mes);
}

/** "Presupuesto 2026.xlsx" + cerrados=8 → "Presupuesto 2026 8+4.xlsx" — mismo patrón
 *  "N+M" que ya usa CAPEX para sus archivos de cierre (N meses cerrados, M restantes en
 *  el año). Usada por `generarArchivoDeCierreOpex` para nombrar la copia nueva. */
export function construirNombreCierre(nombreActual: string, cerrados: number): string {
  const restantes = 12 - cerrados;
  const patronNM = /\d+\s*\+\s*\d+/;
  if (patronNM.test(nombreActual)) {
    return nombreActual.replace(patronNM, `${cerrados}+${restantes}`);
  }
  const punto = nombreActual.lastIndexOf(".");
  const base = punto === -1 ? nombreActual : nombreActual.slice(0, punto);
  const extension = punto === -1 ? "" : nombreActual.slice(punto);
  return `${base} ${cerrados}+${restantes}${extension}`;
}

/**
 * Genera el archivo del siguiente cierre de OPEX: una COPIA exacta del archivo en vivo,
 * con el nombre que sigue en la numeración "N+M" (ej. "8+4" → "9+3") — mismo criterio
 * que "Generar archivo de cierre" en CAPEX. El archivo actual NO se toca (nunca se
 * renombra ni se modifica), así que queda protegido tal como se presentó; el nuevo
 * archivo pasa a ser "el vivo" en cuanto la app lo detecta (mayor N en la carpeta).
 *
 * A diferencia de CAPEX, OPEX no tiene ninguna fórmula de Forecast en el Excel que haga
 * falta reescribir (el Forecast se calcula aparte, en la propia app, a partir de estas
 * mismas celdas) — por eso acá alcanza con copiar el archivo tal cual, sin tocar
 * ninguna celda del archivo nuevo.
 */
export async function generarArchivoDeCierreOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto
): Promise<{ archivo: string }> {
  const mesCierreActual = await leerMesCierre(config, archivo);
  if (mesCierreActual >= 12) {
    throw new ErrorSharePoint("Ya no quedan meses por cerrar este año.");
  }
  const nuevoNombre = construirNombreCierre(archivo.nombre, mesCierreActual + 1);
  if (nuevoNombre.toLowerCase() === archivo.nombre.toLowerCase()) {
    throw new ErrorSharePoint(`El archivo ya se llama "${nuevoNombre}".`);
  }

  const existentes = await listarNombresCarpeta(config, archivo.driveId, archivo.carpetaId);
  if (existentes.some((n) => n.toLowerCase() === nuevoNombre.toLowerCase())) {
    throw new ErrorSharePoint(`Ya existe un archivo llamado "${nuevoNombre}" en esa carpeta.`);
  }

  const contenido = await descargarContenido(config, archivo);
  const nuevoArchivo = await crearArchivo(config, archivo.driveId, archivo.carpetaId, nuevoNombre, contenido);
  return { archivo: nuevoArchivo.nombre };
}
