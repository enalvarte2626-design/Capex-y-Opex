import {
  camposFaltantes,
  camposFaltantesOpex,
  obtenerConfiguracionOpex,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "./sharepoint";
import { leerAnioActivo } from "./anioActivoConfig";

/** Punto de partida si todavía nunca se aprobó ningún año desde el botón "Aprobar y
 *  activar" (o si el Excel no está configurado, ej. en desarrollo local). */
const ANIO_POR_DEFECTO = Number(process.env.ANIO_PRESUPUESTO?.trim()) || 2026;

/**
 * Año de presupuesto activo de CAPEX/OPEX — CADA UNO VIVE EN SU PROPIO EXCEL (archivos
 * distintos), así que se lee por separado. Nunca lanza: si el Excel no responde o no
 * está configurado, cae en ANIO_POR_DEFECTO — esto se usa en el layout de TODA la app,
 * así que una falla de red acá no puede tumbar ninguna página.
 */
export async function anioActivoCapex(): Promise<number> {
  try {
    const config = obtenerConfiguracionSharePoint();
    if (camposFaltantes(config).length > 0) return ANIO_POR_DEFECTO;
    const archivo = await resolverArchivoPorShareUrl(config);
    return await leerAnioActivo(config, archivo, ANIO_POR_DEFECTO);
  } catch {
    return ANIO_POR_DEFECTO;
  }
}

export async function anioActivoOpex(): Promise<number> {
  try {
    const config = obtenerConfiguracionOpex();
    if (camposFaltantesOpex(config).length > 0) return ANIO_POR_DEFECTO;
    const archivo = await resolverArchivoPorShareUrl(config);
    return await leerAnioActivo(config, archivo, ANIO_POR_DEFECTO);
  } catch {
    return ANIO_POR_DEFECTO;
  }
}
