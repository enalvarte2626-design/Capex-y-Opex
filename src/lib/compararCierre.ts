import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import {
  descargarContenido,
  descargarVersionArchivo,
  listarArchivosCarpeta,
  listarVersionesArchivo,
  type VersionArchivo,
} from "./sharepoint";
import { leerWorkbook, extraerProyectos } from "./capex-parse";

const NOMBRES_MES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

/** Umbral para no marcar como "cambio" un ruido de redondeo de centavos. */
const TOLERANCIA = 0.005;

/** "Control Capex Forecast 7+5.xlsm" → 7 (meses cerrados según el nombre del archivo) —
 *  solo un punto de partida sugerido: la persona puede ajustarlo a mano, porque un
 *  ARCHIVO puede seguir editándose después de generarse (como pasó acá: se editó "7+5"
 *  después de presentarlo, antes de generar "8+4" a partir de esa edición). */
function cerradosSugeridosPorNombre(nombre: string): number {
  const m = nombre.match(/(\d+)\s*\+\s*(\d+)/);
  return m ? Number(m[1]) : 0;
}

export interface ArchivoParaComparar {
  nombre: string;
  itemId: string;
  esArchivoActual: boolean;
  cerradosSugeridos: number;
}

/** Lista TODOS los archivos de la misma carpeta que el archivo en vivo (incluido él
 *  mismo) — para elegir contra cuál (o contra qué VERSIÓN anterior del mismo archivo)
 *  comparar. A diferencia de la primera versión de este módulo, ya no se limita a los
 *  que siguen el patrón "N+M": el archivo que hace falta comparar puede ser el mismo
 *  archivo en vivo, en una versión de SharePoint anterior a que alguien lo siguiera
 *  editando después de "cerrarlo". */
export async function listarArchivosParaComparar(
  config: ConfiguracionSharePoint,
  archivoActual: ArchivoResuelto
): Promise<ArchivoParaComparar[]> {
  const archivos = await listarArchivosCarpeta(config, archivoActual.driveId, archivoActual.carpetaId);
  return archivos
    .filter((a) => /\.xlsm$|\.xlsx$/i.test(a.nombre))
    .map((a) => ({
      nombre: a.nombre,
      itemId: a.itemId,
      esArchivoActual: a.itemId === archivoActual.itemId,
      cerradosSugeridos: cerradosSugeridosPorNombre(a.nombre),
    }))
    .sort((a, b) => (a.esArchivoActual === b.esArchivoActual ? 0 : a.esArchivoActual ? -1 : 1));
}

/** Historial de versiones de un archivo puntual — SharePoint guarda una sola cada vez
 *  que alguien lo guarda, sin que nadie la pida a propósito. */
export async function listarVersionesDeArchivo(
  config: ConfiguracionSharePoint,
  driveId: string,
  itemId: string
): Promise<VersionArchivo[]> {
  return listarVersionesArchivo(config, driveId, itemId);
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
  /** true = ese mes ya estaba cerrado (Gasto Real definitivo) en el punto de referencia
   *  elegido — si cambió de todas formas, es justo lo que se quiere detectar: algo que
   *  se movió DESPUÉS de haberse cerrado. */
  eraMesCerrado: boolean;
}

export interface ResumenGrupo {
  grupoNegocio: string;
  /** Diferencia total del grupo en el punto de referencia (Presupuesto Aprobado menos
   *  lo ejecutado/proyectado en ese momento) — negativo = ya se había sobrepasado el
   *  presupuesto; positivo = quedaba margen (ahorro). */
  diferenciaAntes: number;
  diferenciaAhora: number;
  /** diferenciaAhora − diferenciaAntes: positivo = el grupo MEJORÓ (más ahorro o menos
   *  sobrepaso que antes); negativo = el grupo EMPEORÓ desde el punto de referencia. */
  cambio: number;
}

export interface ResultadoComparacion {
  nombreArchivoAnterior: string;
  fechaVersionAnterior: string | null;
  cerradosEnReferencia: number;
  cerradosActual: number;
  proyectosNuevos: string[];
  cambios: CambioLinea[];
  /** Por Grupo de Negocio, ordenado del cambio más grande (en valor absoluto) al más
   *  chico — para ver de un vistazo qué grupo es el que más movió el indicador general. */
  resumenPorGrupo: ResumenGrupo[];
  totalAntes: number;
  totalAhora: number;
  totalCambio: number;
}

/** Presupuesto Aprobado − Gasto Real − Forecast, igual fórmula que ya usa el Dashboard
 *  (resolverProyecto en lib/capex.ts) — acá recalculada aparte porque compararCierre.ts
 *  trabaja con `ProyectoCapex` crudo (sin resolver), y cada snapshot puede tener su
 *  propio "meses cerrados" distinto. */
function diferenciaProyecto(p: { presupuestoAprobado: number; real: number[]; proyectado: number[] }, cerrados: number): number {
  const gastoReal = p.real.reduce((a, b) => a + b, 0);
  const forecast = p.proyectado.slice(cerrados).reduce((a, b) => a + b, 0);
  return p.presupuestoAprobado - gastoReal - forecast;
}

function sumarDiferenciaPorGrupo(
  proyectos: Array<{ grupoNegocio: string; presupuestoAprobado: number; real: number[]; proyectado: number[] }>,
  cerrados: number
): Map<string, number> {
  const mapa = new Map<string, number>();
  for (const p of proyectos) {
    const clave = p.grupoNegocio || "SIN GRUPO";
    mapa.set(clave, (mapa.get(clave) ?? 0) + diferenciaProyecto(p, cerrados));
  }
  return mapa;
}

/**
 * Compara BD_CAPEX en vivo contra un punto de referencia puntual: un archivo distinto
 * (su contenido más reciente) o una VERSIÓN anterior de SharePoint de cualquier archivo
 * (incluido el mismo archivo en vivo) — fila por fila, mismo `filaExcel` en los dos. Un
 * proyecto agregado después del punto de referencia simplemente no tiene con qué
 * compararse, y se reporta aparte en `proyectosNuevos`.
 *
 * `cerradosEnReferencia` es cuántos meses estaban cerrados EN ESE PUNTO DE REFERENCIA
 * (no necesariamente lo que dice el nombre del archivo — la persona lo confirma a mano,
 * porque un archivo puede seguir editándose después de "cerrarse").
 */
export async function compararConReferencia(
  config: ConfiguracionSharePoint,
  archivoActual: ArchivoResuelto,
  hoja: string,
  referencia: { driveId: string; itemId: string; nombre: string; versionId?: string; fechaVersion?: string },
  cerradosEnReferencia: number,
  cerradosActual: number
): Promise<ResultadoComparacion> {
  const [contenidoActual, contenidoAnterior] = await Promise.all([
    descargarContenido(config, archivoActual),
    referencia.versionId
      ? descargarVersionArchivo(config, referencia.driveId, referencia.itemId, referencia.versionId)
      : descargarContenido(config, { driveId: referencia.driveId, itemId: referencia.itemId, nombre: referencia.nombre, carpetaId: "" }),
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
          eraMesCerrado: m < cerradosEnReferencia,
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
          eraMesCerrado: m < cerradosEnReferencia,
        });
      }
    }
  }

  // Los cambios en meses ya cerrados (la causa más probable de lo que se busca) primero.
  cambios.sort((a, b) => (a.eraMesCerrado === b.eraMesCerrado ? 0 : a.eraMesCerrado ? -1 : 1));

  const antesPorGrupo = sumarDiferenciaPorGrupo(proyectosAnteriores, cerradosEnReferencia);
  const ahoraPorGrupo = sumarDiferenciaPorGrupo(proyectosActuales, cerradosActual);
  const gruposTodos = new Set([...antesPorGrupo.keys(), ...ahoraPorGrupo.keys()]);
  const resumenPorGrupo: ResumenGrupo[] = Array.from(gruposTodos)
    .map((grupoNegocio) => {
      const diferenciaAntes = antesPorGrupo.get(grupoNegocio) ?? 0;
      const diferenciaAhora = ahoraPorGrupo.get(grupoNegocio) ?? 0;
      return { grupoNegocio, diferenciaAntes, diferenciaAhora, cambio: diferenciaAhora - diferenciaAntes };
    })
    .sort((a, b) => Math.abs(b.cambio) - Math.abs(a.cambio));

  const totalAntes = resumenPorGrupo.reduce((a, g) => a + g.diferenciaAntes, 0);
  const totalAhora = resumenPorGrupo.reduce((a, g) => a + g.diferenciaAhora, 0);

  return {
    nombreArchivoAnterior: referencia.nombre,
    fechaVersionAnterior: referencia.fechaVersion ?? null,
    cerradosEnReferencia,
    cerradosActual,
    proyectosNuevos,
    cambios,
    resumenPorGrupo,
    totalAntes,
    totalAhora,
    totalCambio: totalAhora - totalAntes,
  };
}
