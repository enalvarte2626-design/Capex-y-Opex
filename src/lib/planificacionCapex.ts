import type { ArchivoResuelto, ConfiguracionSharePoint } from "./sharepoint";
import {
  ErrorSharePoint,
  descargarContenido,
  crearHojaSiNoExiste,
  eliminarHojaSiExiste,
  escribirColumna,
  escribirFila,
  listarHojas,
  renombrarHoja,
} from "./sharepoint";
import { leerWorkbook, extraerProyectos, resolverNombreHoja, ultimaFilaConDatosEscaneada } from "./capex-parse";
import type { ProyectoCapex } from "./capex";
import { escribirAnioActivo } from "./anioActivoConfig";

/** Nombre fijo, para siempre, de la hoja de planificación (borrador) de CAPEX — nunca
 *  cambia de nombre entre años; solo se limpia y se vuelve a llenar cada vez que se
 *  genera un borrador nuevo, y su CONTENIDO reemplaza al de la hoja en vivo al aprobar. */
export const HOJA_PLANIFICACION_CAPEX = "BD_CAPEX (Planificación)";

interface DatosFilaCapexBorrador {
  proyecto: string;
  subNegocio: string;
  grupoNegocio: string;
  detalle: string;
  prioridad: string;
  presupuestoAprobado: number;
  /** Solo se completan al IMPORTAR desde un Excel externo (ver `importarBorradorCapex`)
   *  — al generar el borrador copiando BD_CAPEX en vivo, quedan en su valor por defecto
   *  (en blanco/$0), a propósito, para no arrastrar datos de seguimiento del año pasado. */
  categoria?: string;
  status?: string;
  opex?: string;
  recurso?: string;
  responsable?: string;
  tiempo?: string;
  /** Gasto real/proyectado por mes (índice 0 = enero … 11 = diciembre) — en blanco ($0)
   *  salvo que se importen de un Excel externo. */
  real?: number[];
  proyectado?: number[];
}

/** Escribe una fila de proyecto completa en `hoja`, fila `fila` — mismas columnas/fórmulas
 *  que ya usa BD_CAPEX (A:E datos, F fórmula de Avance, G:AL resto + Presupuesto Aprobado,
 *  AM:AO fórmulas de Gasto Real/Forecast/Diferencia, todas referidas solo a su propia
 *  fila). A propósito NO extiende ninguna fórmula de una fila "TOTAL": la hoja de
 *  planificación no tiene fila TOTAL — ningún indicador de la app depende de ella (todo
 *  se recalcula en la propia app a partir de estas mismas celdas), así que evitamos el
 *  único paso realmente delicado (estirar rangos SUBTOTAL) en una operación que corre
 *  año tras año sin que nadie la pueda revisar después. */
async function escribirFilaCapex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  fila: number,
  datos: DatosFilaCapexBorrador
): Promise<void> {
  const real = datos.real ?? Array(12).fill(0);
  const proyectado = datos.proyectado ?? Array(12).fill(0);
  const meses: number[] = [];
  for (let m = 0; m < 12; m++) meses.push(real[m] ?? 0, proyectado[m] ?? 0);

  // El % de Avance (E) SIEMPRE arranca en 0, incluso al importar de un Excel externo —
  // es un dato de seguimiento propio del año en curso, no algo que tenga sentido
  // arrastrar de un archivo de planificación.
  await escribirFila(config, archivo, hoja, fila, "A", "E", [
    datos.proyecto,
    datos.subNegocio,
    datos.grupoNegocio,
    datos.detalle,
    0,
  ]);
  await escribirColumna(config, archivo, hoja, "F", fila, fila, [
    `=IF(E${fila}=0%,"No iniciado",IF(E${fila}<30%,"Inicio",IF(E${fila}<80%,"En proceso",IF(E${fila}<100%,"Por culminar","Culminado"))))`,
  ]);
  await escribirFila(config, archivo, hoja, fila, "G", "AL", [
    datos.categoria ?? "",
    datos.prioridad,
    datos.status ?? "",
    datos.opex ?? "",
    datos.recurso ?? "",
    datos.responsable ?? "",
    datos.tiempo ?? "",
    ...meses,
    datos.presupuestoAprobado,
  ]);
  await escribirColumna(config, archivo, hoja, "AM", fila, fila, [
    `=SUM(N${fila}+P${fila}+R${fila}+T${fila}+V${fila}+X${fila}+Z${fila}+AB${fila}+AD${fila}+AF${fila}+AH${fila}+AJ${fila})`,
  ]);
  await escribirColumna(config, archivo, hoja, "AN", fila, fila, [
    `=SUM(AC${fila}+AE${fila}+AG${fila}+AI${fila}+AK${fila})`,
  ]);
  await escribirColumna(config, archivo, hoja, "AO", fila, fila, [`=AL${fila}-AM${fila}-AN${fila}`]);
}

/**
 * (Re)genera la hoja de planificación desde cero, copiando la lista de proyectos que
 * hay HOY en la hoja en vivo — en blanco: Presupuesto Aprobado y los 12 meses en $0 (a
 * propósito, para no arrastrar montos del año pasado por error), listos para llenar. Si
 * ya existía un borrador, se borra y se reemplaza por completo — pensado para usarse una
 * vez al empezar a planificar el año, no para "actualizar" sin perder lo ya escrito.
 */
export async function generarBorradorCapex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hojaViva: string
): Promise<{ lineas: number }> {
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const proyectos = extraerProyectos(wb, hojaViva);

  await eliminarHojaSiExiste(config, archivo, HOJA_PLANIFICACION_CAPEX);
  await crearHojaSiNoExiste(config, archivo, HOJA_PLANIFICACION_CAPEX);
  await escribirFila(config, archivo, HOJA_PLANIFICACION_CAPEX, 1, "A", "E", [
    "Proyecto",
    "Sub Negocio",
    "Grupo de Negocio",
    "Detalle",
    "% Avance",
  ]);

  for (let i = 0; i < proyectos.length; i++) {
    const p = proyectos[i];
    await escribirFilaCapex(config, archivo, HOJA_PLANIFICACION_CAPEX, i + 2, {
      proyecto: p.proyecto,
      subNegocio: p.subNegocio,
      grupoNegocio: p.grupoNegocio,
      detalle: p.detalle,
      prioridad: p.prioridad,
      presupuestoAprobado: 0,
    });
  }

  return { lineas: proyectos.length };
}

/** Lee la hoja de planificación (si todavía no existe, devuelve vacío en vez de fallar —
 *  significa que nadie generó el borrador todavía este año). */
export async function leerBorradorCapex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto
): Promise<ProyectoCapex[]> {
  const hojas = await listarHojas(config, archivo);
  if (!hojas.some((h) => h.toLowerCase() === HOJA_PLANIFICACION_CAPEX.toLowerCase())) return [];
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  return extraerProyectos(wb, HOJA_PLANIFICACION_CAPEX);
}

/** Agrega una línea nueva al borrador (además de las copiadas al generarlo) — al final,
 *  sin ningún hueco ni fila TOTAL de por medio. */
export async function agregarLineaBorradorCapex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  datos: DatosFilaCapexBorrador
): Promise<{ fila: number }> {
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const ultimaFila = ultimaFilaConDatosEscaneada(wb, HOJA_PLANIFICACION_CAPEX);
  const filaNueva = Math.max(ultimaFila, 1) + 1;
  await escribirFilaCapex(config, archivo, HOJA_PLANIFICACION_CAPEX, filaNueva, datos);
  return { fila: filaNueva };
}

/**
 * "Aprobar y activar": el borrador pasa a ser la hoja en vivo, y la hoja en vivo saliente
 * queda archivada con el año en el nombre — todo con renombres de hoja (nunca copiando ni
 * borrando filas), así ninguna fórmula corre riesgo de romperse. Después de esto, el
 * borrador ya no existe (se puede volver a generar cuando se quiera planificar el año
 * siguiente) y el año activo de CAPEX sube en 1.
 */
export async function aprobarYActivarCapex(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hojaViva: string,
  anioActivo: number
): Promise<{ anioNuevo: number; hojaArchivada: string }> {
  const hojas = await listarHojas(config, archivo);
  if (!hojas.some((h) => h.toLowerCase() === HOJA_PLANIFICACION_CAPEX.toLowerCase())) {
    throw new ErrorSharePoint('No hay ningún borrador generado todavía — usa "Generar borrador" primero.');
  }
  const contenido = await descargarContenido(config, archivo);
  const wb = leerWorkbook(contenido);
  const lineasBorrador = extraerProyectos(wb, HOJA_PLANIFICACION_CAPEX);
  if (lineasBorrador.length === 0) {
    throw new ErrorSharePoint("El borrador está vacío — agrega al menos una línea antes de aprobar.");
  }

  const hojaArchivada = `${hojaViva} archivo ${anioActivo}`;
  if (hojas.some((h) => h.toLowerCase() === hojaArchivada.toLowerCase())) {
    throw new ErrorSharePoint(
      `Ya existe una hoja "${hojaArchivada}" — parece que este año ya se aprobó antes. Si esto es un error, contacta soporte antes de reintentar.`
    );
  }

  // Confirma que la hoja en vivo todavía existe con ese nombre exacto antes de tocar nada
  // (defensivo: si alguien la renombró a mano desde Excel, mejor fallar claro que dejar
  // el archivo en un estado raro a la mitad).
  if (!hojas.some((h) => h.toLowerCase() === hojaViva.toLowerCase())) {
    throw new ErrorSharePoint(`No se encontró la hoja en vivo "${hojaViva}".`);
  }

  await renombrarHoja(config, archivo, hojaViva, hojaArchivada);
  await renombrarHoja(config, archivo, HOJA_PLANIFICACION_CAPEX, hojaViva);
  const anioNuevo = anioActivo + 1;
  await escribirAnioActivo(config, archivo, anioNuevo);

  return { anioNuevo, hojaArchivada };
}

/**
 * Reemplaza TODO el borrador de planificación por lo que traiga un Excel subido a mano
 * (mismo layout de columnas que BD_CAPEX en vivo, incluido Real/Proyectado mes a mes) —
 * a diferencia de `generarBorradorCapex` (que copia los proyectos ya existentes en
 * blanco) y de `agregarLineaBorradorCapex` (que suma una línea a lo que ya había), esto
 * BORRA el borrador actual y lo vuelve a llenar solo con lo que trae el archivo.
 * Busca la hoja `hojaViva` (ej. "BD_CAPEX") dentro del Excel subido; si no la
 * encuentra con ese nombre, usa la primera hoja del archivo.
 */
export async function importarBorradorCapex(
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
  const proyectos = extraerProyectos(wbSubido, hojaLeida);
  if (proyectos.length === 0) {
    throw new ErrorSharePoint(`No se encontró ningún proyecto en la hoja "${hojaLeida}" del Excel subido.`);
  }

  await eliminarHojaSiExiste(config, archivo, HOJA_PLANIFICACION_CAPEX);
  await crearHojaSiNoExiste(config, archivo, HOJA_PLANIFICACION_CAPEX);
  await escribirFila(config, archivo, HOJA_PLANIFICACION_CAPEX, 1, "A", "E", [
    "Proyecto",
    "Sub Negocio",
    "Grupo de Negocio",
    "Detalle",
    "% Avance",
  ]);

  for (let i = 0; i < proyectos.length; i++) {
    const p = proyectos[i];
    await escribirFilaCapex(config, archivo, HOJA_PLANIFICACION_CAPEX, i + 2, {
      proyecto: p.proyecto,
      subNegocio: p.subNegocio,
      grupoNegocio: p.grupoNegocio,
      detalle: p.detalle,
      prioridad: p.prioridad,
      presupuestoAprobado: p.presupuestoAprobado,
      categoria: p.categoria,
      status: p.status,
      opex: p.opex,
      recurso: p.recurso,
      responsable: p.responsable,
      tiempo: p.tiempo,
      real: p.real,
      proyectado: p.proyectado,
    });
  }

  return { lineas: proyectos.length, hojaLeida };
}
