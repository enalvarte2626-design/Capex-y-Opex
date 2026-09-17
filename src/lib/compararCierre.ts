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

const RANGO_DIACRITICOS = /[̀-ͯ]/g;

/** Clave de identidad de un proyecto: Proyecto + Detalle, normalizado — el mismo
 *  criterio que ya usa el resto de la app para cruzar BD_CAPEX contra otra hoja (ver
 *  normalizarNombreProyecto en app/page.tsx). NO se usa el número de fila para
 *  identificar un proyecto entre dos archivos: si alguien inserta una fila, borra una,
 *  u ordena la hoja en Excel entre un cierre y el siguiente, el mismo número de fila
 *  puede terminar apuntando a un proyecto totalmente distinto, y comparar "fila 15 de
 *  antes" contra "fila 15 de ahora" mezclaría dos proyectos sin que nadie se entere. */
function claveProyecto(p: { proyecto: string; detalle: string }): string {
  const normalizar = (t: string) => t.trim().toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
  return `${normalizar(p.proyecto)}|${normalizar(p.detalle)}`;
}

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
  presupuestoAprobadoAntes: number;
  presupuestoAprobadoAhora: number;
  forecastAntes: number;
  forecastAhora: number;
  /** Diferencia total del grupo en el punto de referencia (Presupuesto Aprobado menos
   *  lo ejecutado/proyectado en ese momento) — negativo = ya se había sobrepasado el
   *  presupuesto; positivo = quedaba margen (ahorro). */
  diferenciaAntes: number;
  diferenciaAhora: number;
  /** diferenciaAhora − diferenciaAntes: positivo = el grupo MEJORÓ (más ahorro o menos
   *  sobrepaso que antes); negativo = el grupo EMPEORÓ desde el punto de referencia. */
  cambio: number;
}

export interface ProyectoNuevo {
  proyecto: string;
  detalle: string;
  grupoNegocio: string;
  presupuestoAprobado: number;
  forecast: number;
}

export interface ResultadoComparacion {
  nombreArchivoAnterior: string;
  fechaVersionAnterior: string | null;
  cerradosEnReferencia: number;
  cerradosActual: number;
  /** Proyectos que hoy existen pero no estaban todavía en el punto de referencia — no
   *  hay "antes" con qué compararlos, así que no generan ninguna fila en `cambios`, pero
   *  sí están contados en los totales de `resumenPorGrupo` (su grupo, con estos mismos
   *  valores). Se listan con su Presupuesto Aprobado y Forecast actuales para poder
   *  revisar qué se les cargó, ya que no aparecen en ningún otro lado de esta pantalla. */
  proyectosNuevos: ProyectoNuevo[];
  cambios: CambioLinea[];
  /** Por Grupo de Negocio, ordenado del cambio más grande (en valor absoluto) al más
   *  chico — para ver de un vistazo qué grupo es el que más movió el indicador general. */
  resumenPorGrupo: ResumenGrupo[];
  totalAntes: number;
  totalAhora: number;
  totalCambio: number;
}

interface TotalesProyecto {
  presupuestoAprobado: number;
  gastoReal: number;
  forecast: number;
  diferencia: number;
}

/** Presupuesto Aprobado, Gasto Real, Forecast y Diferencia — misma fórmula que ya usa el
 *  Dashboard (resolverProyecto en lib/capex.ts) — recalculada aparte porque
 *  compararCierre.ts trabaja con `ProyectoCapex` crudo (sin resolver), y cada snapshot
 *  puede tener su propio "meses cerrados" distinto. */
function totalesProyecto(p: { presupuestoAprobado: number; real: number[]; proyectado: number[] }, cerrados: number): TotalesProyecto {
  const gastoReal = p.real.reduce((a, b) => a + b, 0);
  const forecast = p.proyectado.slice(cerrados).reduce((a, b) => a + b, 0);
  return { presupuestoAprobado: p.presupuestoAprobado, gastoReal, forecast, diferencia: p.presupuestoAprobado - gastoReal - forecast };
}

function sumarPorGrupo(
  proyectos: Array<{ grupoNegocio: string; presupuestoAprobado: number; real: number[]; proyectado: number[] }>,
  cerrados: number
): Map<string, TotalesProyecto> {
  const mapa = new Map<string, TotalesProyecto>();
  for (const p of proyectos) {
    const clave = p.grupoNegocio || "SIN GRUPO";
    const t = totalesProyecto(p, cerrados);
    const acc = mapa.get(clave) ?? { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 };
    mapa.set(clave, {
      presupuestoAprobado: acc.presupuestoAprobado + t.presupuestoAprobado,
      gastoReal: acc.gastoReal + t.gastoReal,
      forecast: acc.forecast + t.forecast,
      diferencia: acc.diferencia + t.diferencia,
    });
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
  cerradosActual: number,
  /** Si viene, solo se comparan los proyectos cuya Prioridad (tal cual el texto en
   *  Excel, ej. "1", "2") esté en esta lista — el resto queda completamente afuera de
   *  cambios, proyectosNuevos y el resumen por grupo, no solo escondido en la pantalla. */
  prioridades?: string[]
): Promise<ResultadoComparacion> {
  const [contenidoActual, contenidoAnterior] = await Promise.all([
    descargarContenido(config, archivoActual),
    referencia.versionId
      ? descargarVersionArchivo(config, referencia.driveId, referencia.itemId, referencia.versionId)
      : descargarContenido(config, { driveId: referencia.driveId, itemId: referencia.itemId, nombre: referencia.nombre, carpetaId: "" }),
  ]);

  const filtroPrioridad = prioridades && prioridades.length > 0 ? new Set(prioridades.map((p) => p.trim())) : null;
  const filtrarPrioridad = <T extends { prioridad: string }>(lista: T[]): T[] =>
    filtroPrioridad ? lista.filter((p) => filtroPrioridad.has(p.prioridad.trim())) : lista;

  const proyectosActuales = filtrarPrioridad(extraerProyectos(leerWorkbook(contenidoActual), hoja));
  const proyectosAnteriores = filtrarPrioridad(extraerProyectos(leerWorkbook(contenidoAnterior), hoja));

  // Mapa por Proyecto+Detalle — ver el comentario de claveProyecto sobre por qué NO se
  // usa el número de fila como primer criterio. Si dos proyectos de "antes" comparten la
  // misma clave (nombre y detalle idénticos, algo raro pero posible), se marca como
  // ambigua (null) para NO arriesgarse a cruzar dos proyectos distintos por error.
  const mapaAnterior = new Map<string, (typeof proyectosAnteriores)[number] | null>();
  for (const p of proyectosAnteriores) {
    const clave = claveProyecto(p);
    mapaAnterior.set(clave, mapaAnterior.has(clave) ? null : p);
  }
  // Respaldo por número de fila — para cuando el proyecto SÍ es el mismo pero se le
  // corrigió el nombre o el detalle entre un cierre y el siguiente (un cambio de texto
  // legítimo no debería hacer que un proyecto que "siempre existió" salga marcado como
  // nuevo). Solo se usa si el emparejamiento por texto no encontró nada, y solo una vez
  // por fila (`filasAnterioresUsadas` evita que la misma fila de "antes" quede
  // reclamada dos veces si además hubo una reordenada real de por medio).
  const mapaAnteriorPorFila = new Map(proyectosAnteriores.map((p) => [p.filaExcel, p]));
  const filasAnterioresUsadas = new Set<number>();
  for (const p of proyectosActuales) {
    const porTexto = mapaAnterior.get(claveProyecto(p));
    if (porTexto) filasAnterioresUsadas.add(porTexto.filaExcel);
  }

  const cambios: CambioLinea[] = [];
  const proyectosNuevos: ProyectoNuevo[] = [];

  for (const actual of proyectosActuales) {
    let anterior = mapaAnterior.get(claveProyecto(actual));
    if (!anterior) {
      const porFila = mapaAnteriorPorFila.get(actual.filaExcel);
      if (porFila && !filasAnterioresUsadas.has(porFila.filaExcel)) {
        anterior = porFila;
        filasAnterioresUsadas.add(porFila.filaExcel);
      }
    }
    if (!anterior) {
      const t = totalesProyecto(actual, cerradosActual);
      proyectosNuevos.push({
        proyecto: actual.proyecto,
        detalle: actual.detalle,
        grupoNegocio: actual.grupoNegocio,
        presupuestoAprobado: t.presupuestoAprobado,
        forecast: t.forecast,
      });
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

  const antesPorGrupo = sumarPorGrupo(proyectosAnteriores, cerradosEnReferencia);
  const ahoraPorGrupo = sumarPorGrupo(proyectosActuales, cerradosActual);
  const gruposTodos = new Set([...antesPorGrupo.keys(), ...ahoraPorGrupo.keys()]);
  const vacio: TotalesProyecto = { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 };
  const resumenPorGrupo: ResumenGrupo[] = Array.from(gruposTodos)
    .map((grupoNegocio) => {
      const antes = antesPorGrupo.get(grupoNegocio) ?? vacio;
      const ahora = ahoraPorGrupo.get(grupoNegocio) ?? vacio;
      return {
        grupoNegocio,
        presupuestoAprobadoAntes: antes.presupuestoAprobado,
        presupuestoAprobadoAhora: ahora.presupuestoAprobado,
        forecastAntes: antes.forecast,
        forecastAhora: ahora.forecast,
        diferenciaAntes: antes.diferencia,
        diferenciaAhora: ahora.diferencia,
        cambio: ahora.diferencia - antes.diferencia,
      };
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
