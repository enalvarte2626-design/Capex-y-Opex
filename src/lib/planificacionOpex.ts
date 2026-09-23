import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import {
  ErrorSharePoint,
  descargarContenido,
  crearHojaSiNoExiste,
  eliminarHojaSiExiste,
  escribirFila,
  listarHojas,
  renombrarHoja,
} from "./sharepoint";
import { leerWorkbook, resolverNombreHoja, ultimaFilaConDatosEscaneada } from "./capex-parse";
import { columnaALetra } from "./capex-editable";
import { COL_PPTO_OPEX, extraerPresupuestoOpex } from "./opex-parse";
import type { ProyectoCapex } from "./capex";
import { escribirAnioActivo } from "./anioActivoConfig";

/** Nombre fijo, para siempre, de la hoja de planificación (borrador) de OPEX — mismo
 *  criterio que HOJA_PLANIFICACION_CAPEX en planificacionCapex.ts. */
export const HOJA_PLANIFICACION_OPEX = "Presupuesto (Planificación)";

/** Última columna que se escribe en una fila (AI = Presupuesto Aprobado, índice 34). */
const ULTIMA_COLUMNA = columnaALetra(COL_PPTO_OPEX.presupuestoAprobado);

interface DatosLineaOpexBorrador {
  empresa: string;
  grupoGasto: string;
  subgrupoGasto: string;
  lineaGasto: string;
  moneda: string;
  detalle: string;
  responsable: string;
  presupuestoAprobado: number;
  /** Solo se completan al IMPORTAR desde un Excel externo (ver `importarBorradorOpex`)
   *  — al generar el borrador copiando el vivo, quedan en $0 a propósito. */
  status?: string;
  real?: number[];
  proyectado?: number[];
}

/** Arma la fila completa (A:AI) para una línea de OPEX — "Presupuesto 2026" se lee por
 *  posición de columna, no como Tabla de Excel (ver COL_PPTO_OPEX), así que no hace
 *  falta ninguna Tabla real ni fórmula de por medio: solo escribir los valores en su
 *  columna correcta. Los 24 meses (Proyectado/Real alternados) arrancan en 0, salvo que
 *  se importen de un Excel externo. */
function filaOpex(datos: DatosLineaOpexBorrador): Array<string | number> {
  const fila: Array<string | number> = new Array(COL_PPTO_OPEX.presupuestoAprobado + 1).fill("");
  fila[COL_PPTO_OPEX.empresa] = datos.empresa;
  fila[COL_PPTO_OPEX.grupoGasto] = datos.grupoGasto;
  fila[COL_PPTO_OPEX.subgrupoGasto] = datos.subgrupoGasto;
  fila[COL_PPTO_OPEX.lineaGasto] = datos.lineaGasto;
  fila[COL_PPTO_OPEX.status] = datos.status || "Activa";
  fila[COL_PPTO_OPEX.moneda] = datos.moneda || "USD";
  fila[COL_PPTO_OPEX.detalle] = datos.detalle;
  fila[COL_PPTO_OPEX.responsable] = datos.responsable;
  const real = datos.real ?? Array(12).fill(0);
  const proyectado = datos.proyectado ?? Array(12).fill(0);
  for (let m = 0; m < 12; m++) {
    fila[COL_PPTO_OPEX.primerMesProyectado + m * 2] = proyectado[m] ?? 0;
    fila[COL_PPTO_OPEX.primerMesReal + m * 2] = real[m] ?? 0;
  }
  fila[COL_PPTO_OPEX.presupuestoAprobado] = datos.presupuestoAprobado;
  return fila;
}

/**
 * (Re)genera la hoja de planificación desde cero, copiando la lista de líneas que hay
 * HOY en la hoja en vivo — en blanco: Presupuesto Aprobado y los 12 meses en $0. Mismo
 * criterio que `generarBorradorCapex`.
 */
export async function generarBorradorOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hojaViva: string
): Promise<{ lineas: number }> {
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const lineas = extraerPresupuestoOpex(wb, hojaViva);

  await eliminarHojaSiExiste(config, archivo, HOJA_PLANIFICACION_OPEX);
  await crearHojaSiNoExiste(config, archivo, HOJA_PLANIFICACION_OPEX);
  const encabezado: Array<string | number> = new Array(COL_PPTO_OPEX.presupuestoAprobado + 1).fill("");
  encabezado[COL_PPTO_OPEX.empresa] = "Empresa";
  encabezado[COL_PPTO_OPEX.grupoGasto] = "Grupo de Gasto";
  encabezado[COL_PPTO_OPEX.subgrupoGasto] = "Subgrupo de Gasto";
  encabezado[COL_PPTO_OPEX.lineaGasto] = "Línea de Gasto";
  encabezado[COL_PPTO_OPEX.status] = "Status";
  encabezado[COL_PPTO_OPEX.moneda] = "Moneda";
  encabezado[COL_PPTO_OPEX.detalle] = "Detalle";
  encabezado[COL_PPTO_OPEX.responsable] = "Responsable";
  encabezado[COL_PPTO_OPEX.presupuestoAprobado] = "Ppto Aprobado";
  await escribirFila(config, archivo, HOJA_PLANIFICACION_OPEX, 1, "A", ULTIMA_COLUMNA, encabezado);

  for (let i = 0; i < lineas.length; i++) {
    const l = lineas[i];
    const fila = filaOpex({
      empresa: l.subNegocio,
      grupoGasto: l.grupoNegocio,
      subgrupoGasto: l.prioridad,
      lineaGasto: l.proyecto,
      moneda: "USD",
      detalle: l.detalle,
      responsable: l.responsable,
      presupuestoAprobado: 0,
    });
    await escribirFila(config, archivo, HOJA_PLANIFICACION_OPEX, i + 2, "A", ULTIMA_COLUMNA, fila);
  }

  return { lineas: lineas.length };
}

/** Lee la hoja de planificación (si todavía no existe, devuelve vacío). */
export async function leerBorradorOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto
): Promise<ProyectoCapex[]> {
  const hojas = await listarHojas(config, archivo);
  if (!hojas.some((h) => h.toLowerCase() === HOJA_PLANIFICACION_OPEX.toLowerCase())) return [];
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  return extraerPresupuestoOpex(wb, HOJA_PLANIFICACION_OPEX);
}

/** Agrega una línea nueva al borrador — al final, sin ningún hueco. */
export async function agregarLineaBorradorOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  datos: DatosLineaOpexBorrador
): Promise<{ fila: number }> {
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const ultimaFila = ultimaFilaConDatosEscaneada(wb, HOJA_PLANIFICACION_OPEX);
  const filaNueva = Math.max(ultimaFila, 1) + 1;
  await escribirFila(config, archivo, HOJA_PLANIFICACION_OPEX, filaNueva, "A", ULTIMA_COLUMNA, filaOpex(datos));
  return { fila: filaNueva };
}

/** "Aprobar y activar" para OPEX — mismo criterio que `aprobarYActivarCapex` (renombres
 *  de hoja, nunca copia ni borra filas). */
export async function aprobarYActivarOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hojaViva: string,
  anioActivo: number
): Promise<{ anioNuevo: number; hojaArchivada: string }> {
  const hojas = await listarHojas(config, archivo);
  if (!hojas.some((h) => h.toLowerCase() === HOJA_PLANIFICACION_OPEX.toLowerCase())) {
    throw new ErrorSharePoint('No hay ningún borrador generado todavía — usa "Generar borrador" primero.');
  }
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const lineasBorrador = extraerPresupuestoOpex(wb, HOJA_PLANIFICACION_OPEX);
  if (lineasBorrador.length === 0) {
    throw new ErrorSharePoint("El borrador está vacío — agrega al menos una línea antes de aprobar.");
  }

  const hojaArchivada = `${hojaViva} archivo ${anioActivo}`;
  if (hojas.some((h) => h.toLowerCase() === hojaArchivada.toLowerCase())) {
    throw new ErrorSharePoint(
      `Ya existe una hoja "${hojaArchivada}" — parece que este año ya se aprobó antes. Si esto es un error, contacta soporte antes de reintentar.`
    );
  }
  if (!hojas.some((h) => h.toLowerCase() === hojaViva.toLowerCase())) {
    throw new ErrorSharePoint(`No se encontró la hoja en vivo "${hojaViva}".`);
  }

  await renombrarHoja(config, archivo, hojaViva, hojaArchivada);
  await renombrarHoja(config, archivo, HOJA_PLANIFICACION_OPEX, hojaViva);
  const anioNuevo = anioActivo + 1;
  await escribirAnioActivo(config, archivo, anioNuevo);

  return { anioNuevo, hojaArchivada };
}

/**
 * Reemplaza TODO el borrador de planificación de OPEX por lo que traiga un Excel subido
 * a mano (mismo layout de columnas que "Presupuesto 2026" en vivo, incluido
 * Proyectado/Real mes a mes) — mismo criterio que `importarBorradorCapex`. Busca la
 * hoja `hojaViva` dentro del Excel subido; si no la encuentra con ese nombre, usa la
 * primera hoja del archivo.
 */
export async function importarBorradorOpex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hojaViva: string,
  contenidoExcelSubido: Buffer
): Promise<{ lineas: number; hojaLeida: string }> {
  const wbSubido = leerWorkbook(contenidoExcelSubido);
  const nombreResuelto = resolverNombreHoja(wbSubido, hojaViva);
  const hojaLeida = wbSubido.Sheets[nombreResuelto] ? nombreResuelto : wbSubido.SheetNames[0];
  if (!hojaLeida) {
    throw new ErrorSharePoint("El Excel subido no tiene ninguna hoja con datos.");
  }
  const lineas = extraerPresupuestoOpex(wbSubido, hojaLeida);
  if (lineas.length === 0) {
    throw new ErrorSharePoint(`No se encontró ninguna línea de gasto en la hoja "${hojaLeida}" del Excel subido.`);
  }

  await eliminarHojaSiExiste(config, archivo, HOJA_PLANIFICACION_OPEX);
  await crearHojaSiNoExiste(config, archivo, HOJA_PLANIFICACION_OPEX);
  const encabezado: Array<string | number> = new Array(COL_PPTO_OPEX.presupuestoAprobado + 1).fill("");
  encabezado[COL_PPTO_OPEX.empresa] = "Empresa";
  encabezado[COL_PPTO_OPEX.grupoGasto] = "Grupo de Gasto";
  encabezado[COL_PPTO_OPEX.subgrupoGasto] = "Subgrupo de Gasto";
  encabezado[COL_PPTO_OPEX.lineaGasto] = "Línea de Gasto";
  encabezado[COL_PPTO_OPEX.status] = "Status";
  encabezado[COL_PPTO_OPEX.moneda] = "Moneda";
  encabezado[COL_PPTO_OPEX.detalle] = "Detalle";
  encabezado[COL_PPTO_OPEX.responsable] = "Responsable";
  encabezado[COL_PPTO_OPEX.presupuestoAprobado] = "Ppto Aprobado";
  await escribirFila(config, archivo, HOJA_PLANIFICACION_OPEX, 1, "A", ULTIMA_COLUMNA, encabezado);

  for (let i = 0; i < lineas.length; i++) {
    const l = lineas[i];
    const fila = filaOpex({
      empresa: l.subNegocio,
      grupoGasto: l.grupoNegocio,
      subgrupoGasto: l.prioridad,
      lineaGasto: l.proyecto,
      moneda: "USD",
      detalle: l.detalle,
      responsable: l.responsable,
      presupuestoAprobado: l.presupuestoAprobado,
      status: l.status,
      real: l.real,
      proyectado: l.proyectado,
    });
    await escribirFila(config, archivo, HOJA_PLANIFICACION_OPEX, i + 2, "A", ULTIMA_COLUMNA, fila);
  }

  return { lineas: lineas.length, hojaLeida };
}
