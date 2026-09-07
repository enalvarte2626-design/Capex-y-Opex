import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import { crearHojaSiNoExiste, escribirCelda, leerCelda } from "./sharepoint";
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
 *  mes desde el botón todavía), usa MES_CIERRE_POR_DEFECTO como punto de partida. */
export async function leerMesCierre(config: ConfiguracionSharePoint, archivo: ArchivoResuelto): Promise<number> {
  try {
    const valor = await leerCelda(config, archivo, HOJA_CONFIG, CELDA_VALOR);
    if (Number.isInteger(valor) && valor >= 1 && valor <= 12) return valor;
  } catch {
    // La hoja "Config App" todavía no existe — nadie cerró ningún mes desde el botón.
  }
  return MES_CIERRE_POR_DEFECTO;
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
