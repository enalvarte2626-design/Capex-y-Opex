import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import { descargarContenido, listarNombresCarpeta, resolverArchivoPorNombreEnCarpeta } from "./sharepoint";
import { leerWorkbook, extraerProyectos } from "./capex-parse";

const NOMBRES_MES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

/** Umbral para no marcar como "cambio" un ruido de redondeo de centavos. */
const TOLERANCIA = 0.005;

/** "Control Capex Forecast 7+5.xlsm" → { cerrados: 7, restantes: 5 }. Mismo patrón que
 *  ya usa /api/capex/cerrar-mes para generar el siguiente archivo. */
function leerPatronNombre(nombre: string): { cerrados: number; restantes: number } | null {
  const m = nombre.match(/(\d+)\s*\+\s*(\d+)/);
  if (!m) return null;
  return { cerrados: Number(m[1]), restantes: Number(m[2]) };
}

export interface ArchivoCierreDisponible {
  nombre: string;
  cerrados: number;
}

/** Lista los archivos de cierre que quedaron guardados en la misma carpeta que el
 *  archivo en vivo — cada vez que se usa "Generar archivo de cierre" queda ahí una
 *  copia exacta de cómo estaba todo justo antes de cerrar ese mes, sin tocar nada. Se
 *  excluye el archivo actual (no tiene sentido "compararlo contra sí mismo"), y se
 *  ordena del cierre más reciente al más antiguo. */
export async function listarArchivosCierreDisponibles(
  config: ConfiguracionSharePoint,
  archivoActual: ArchivoResuelto
): Promise<ArchivoCierreDisponible[]> {
  const nombres = await listarNombresCarpeta(config, archivoActual.driveId, archivoActual.carpetaId);
  const disponibles: ArchivoCierreDisponible[] = [];
  for (const nombre of nombres) {
    if (nombre.toLowerCase() === archivoActual.nombre.toLowerCase()) continue;
    const patron = leerPatronNombre(nombre);
    if (!patron) continue; // no sigue el patrón "N+M" — no es un archivo de cierre de este flujo
    disponibles.push({ nombre, cerrados: patron.cerrados });
  }
  return disponibles.sort((a, b) => b.cerrados - a.cerrados);
}

export interface CambioLinea {
  filaExcel: number;
  proyecto: string;
  detalle: string;
  grupoNegocio: string;
  campo: string;
  valorAnterior: number;
  valorActual: number;
  diferencia: number;
  /** true = ese mes ya estaba cerrado (Gasto Real definitivo) en el archivo anterior —
   *  si cambió de todas formas, es justo lo que se quiere detectar: algo que se movió
   *  DESPUÉS de haberse cerrado. */
  eraMesCerrado: boolean;
}

export interface ResultadoComparacion {
  archivoAnterior: string;
  cerradosEnArchivoAnterior: number;
  proyectosNuevos: string[]; // proyectos que no existían todavía en el archivo anterior
  cambios: CambioLinea[];
}

/**
 * Compara la hoja BD_CAPEX actual contra la de un archivo de cierre anterior, fila por
 * fila (mismo `filaExcel` en los dos — un proyecto agregado después del cierre anterior
 * simplemente no tiene con qué compararse, y se reporta aparte en `proyectosNuevos`).
 * Reporta cada celda que cambió: Presupuesto Aprobado, y cada uno de los 24 meses
 * (Real/Proyectado) — marcando en especial los meses que YA estaban cerrados en el
 * archivo anterior, porque un cambio ahí es la causa más probable de que un indicador
 * ya cerrado "salga distinto" en este cierre.
 */
export async function compararConCierreAnterior(
  config: ConfiguracionSharePoint,
  archivoActual: ArchivoResuelto,
  hoja: string,
  nombreArchivoAnterior: string
): Promise<ResultadoComparacion> {
  const patron = leerPatronNombre(nombreArchivoAnterior);
  const cerradosEnArchivoAnterior = patron?.cerrados ?? 0;

  const archivoAnterior = await resolverArchivoPorNombreEnCarpeta(
    config,
    archivoActual.driveId,
    archivoActual.carpetaId,
    nombreArchivoAnterior
  );

  const [contenidoActual, contenidoAnterior] = await Promise.all([
    descargarContenido(config, archivoActual),
    descargarContenido(config, archivoAnterior),
  ]);

  const proyectosActuales = extraerProyectos(leerWorkbook(contenidoActual), hoja);
  const proyectosAnteriores = extraerProyectos(leerWorkbook(contenidoAnterior), hoja);
  const mapaAnterior = new Map(proyectosAnteriores.map((p) => [p.filaExcel, p]));

  const cambios: CambioLinea[] = [];
  const proyectosNuevos: string[] = [];

  for (const actual of proyectosActuales) {
    const anterior = mapaAnterior.get(actual.filaExcel);
    if (!anterior) {
      proyectosNuevos.push(actual.detalle ? `${actual.proyecto} — ${actual.detalle}` : actual.proyecto);
      continue;
    }

    const etiqueta = { proyecto: actual.proyecto, detalle: actual.detalle, grupoNegocio: actual.grupoNegocio };

    if (Math.abs(actual.presupuestoAprobado - anterior.presupuestoAprobado) > TOLERANCIA) {
      cambios.push({
        filaExcel: actual.filaExcel,
        ...etiqueta,
        campo: "Presupuesto Aprobado",
        valorAnterior: anterior.presupuestoAprobado,
        valorActual: actual.presupuestoAprobado,
        diferencia: actual.presupuestoAprobado - anterior.presupuestoAprobado,
        eraMesCerrado: false,
      });
    }

    for (let m = 0; m < 12; m++) {
      if (Math.abs(actual.real[m] - anterior.real[m]) > TOLERANCIA) {
        cambios.push({
          filaExcel: actual.filaExcel,
          ...etiqueta,
          campo: `Gasto Real de ${NOMBRES_MES[m]}`,
          valorAnterior: anterior.real[m],
          valorActual: actual.real[m],
          diferencia: actual.real[m] - anterior.real[m],
          eraMesCerrado: m < cerradosEnArchivoAnterior,
        });
      }
      if (Math.abs(actual.proyectado[m] - anterior.proyectado[m]) > TOLERANCIA) {
        cambios.push({
          filaExcel: actual.filaExcel,
          ...etiqueta,
          campo: `Gasto Proyectado de ${NOMBRES_MES[m]}`,
          valorAnterior: anterior.proyectado[m],
          valorActual: actual.proyectado[m],
          diferencia: actual.proyectado[m] - anterior.proyectado[m],
          eraMesCerrado: m < cerradosEnArchivoAnterior,
        });
      }
    }
  }

  // Los cambios en meses ya cerrados (la causa más probable de lo que se busca) primero.
  cambios.sort((a, b) => (a.eraMesCerrado === b.eraMesCerrado ? 0 : a.eraMesCerrado ? -1 : 1));

  return { archivoAnterior: nombreArchivoAnterior, cerradosEnArchivoAnterior, proyectosNuevos, cambios };
}
