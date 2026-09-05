import * as XLSX from "xlsx";
import type { FacturaCapex, ItemConMeses, ProyectoCapex } from "./capex";

/** Índices de columna (0-based) dentro de "Control de Facturas-Capex 25fEB". */
export const COL_FACTURAS = {
  periodoFacturado: 0, // A
  recurso: 1, // B
  proveedor: 2, // C
  responsable: 3, // D: Resp.
  proyecto: 4, // E
  monto: 5, // F: Monto Final (sin IGV)
  numeroFactura: 6, // G: N° Factura
  registrado: 7, // H
  comentarios: 8, // I
} as const;

/** Índices de columna (0-based) dentro de BD_CAPEX, según el layout confirmado del archivo. */
export const COL_BD = {
  proyecto: 0,
  subNegocio: 1,
  grupoNegocio: 2,
  detalle: 3,
  avancePct: 4, // E: %
  avance: 5, // F: Avance
  categoria: 6,
  prioridad: 7,
  status: 8,
  opex: 9, // J: OPEX
  recurso: 10, // K: Recurso
  responsable: 11,
  tiempo: 12, // M: TIEMPO
  primerMesReal: 13, // N: Gasto Real Enero-26 (luego alternan Real, Proyectado x 12 meses)
  presupuestoAprobado: 37, // AL
} as const;

/** Índices de columna (0-based) dentro de "Proyeccion 2026_vm". Encabezado en la fila 2 (índice 1). */
const COL_VM = {
  proyecto: 0,
  grupoNegocio: 2,
  // Cada "Proyecto" trae varias filas — una por Detalle (ej. una por sub-tarea/línea de
  // costo), y cada una puede tener su propia Prioridad. No son filas duplicadas: hay que
  // cruzarlas contra BD_CAPEX por Proyecto+Detalle, nunca agrupar solo por Proyecto.
  detalle: 3,
  prioridad: 5,
  primerMes: 12, // M: Gasto Proyectado Enero-26 … 12 columnas seguidas hasta Diciembre (X)
  totalUsd: 26, // AA: TOTAL USD (fuente de verdad en dólares; el total en soles se deriva de este)
} as const;

const FILA_ENCABEZADO_VM = 1; // 0-based: la fila 2 del Excel

function aNumero(valor: unknown): number {
  if (typeof valor === "number") return valor;
  if (valor == null) return 0;
  const texto = String(valor).trim();
  if (!texto || texto === "-") return 0;
  const negativo = /^\(.*\)$/.test(texto) || texto.startsWith("-");
  const limpio = texto.replace(/[^0-9.,-]/g, "").replace(/,/g, "");
  const numero = parseFloat(limpio);
  if (Number.isNaN(numero)) return 0;
  return negativo ? -Math.abs(numero) : numero;
}

function aTexto(valor: unknown): string {
  return String(valor ?? "").trim();
}

/** Lee el workbook una sola vez; se reutiliza para extraer varias hojas sin descargar de nuevo. */
export function leerWorkbook(bufer: Buffer): XLSX.WorkBook {
  return XLSX.read(bufer, { type: "buffer", cellDates: true });
}

function filasDeHoja(wb: XLSX.WorkBook, nombreHoja: string): unknown[][] {
  const hoja = wb.Sheets[nombreHoja];
  if (!hoja) {
    const disponibles = wb.SheetNames.join(", ");
    throw new Error(`No se encontró la hoja "${nombreHoja}" en el archivo. Hojas disponibles: ${disponibles}.`);
  }
  return XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true, defval: "" });
}

/** Última fila (1-based) con datos en la hoja, según el rango usado del Excel. */
export function ultimaFilaConDatos(wb: XLSX.WorkBook, nombreHoja: string): number {
  const hoja = wb.Sheets[nombreHoja];
  if (!hoja || !hoja["!ref"]) return 1;
  const rango = XLSX.utils.decode_range(hoja["!ref"]);
  return rango.e.r + 1; // decode_range es 0-based
}

/**
 * Última fila (1-based) con datos reales, escaneando celda por celda — a diferencia de
 * `ultimaFilaConDatos`, no confía en "!ref" (en hojas donde alguna vez se seleccionó o
 * dio formato a toda la columna, "!ref" queda marcando la última fila de la hoja
 * completa, no la última con contenido real).
 */
export function ultimaFilaConDatosEscaneada(wb: XLSX.WorkBook, nombreHoja: string): number {
  const filas = filasDeHoja(wb, nombreHoja);
  let ultima = 0;
  for (let i = 0; i < filas.length; i++) {
    if (filas[i].some((c) => String(c ?? "").trim() !== "")) ultima = i + 1;
  }
  return ultima;
}

/** Convierte una fecha a serial de Excel (días desde 1899-12-30), como ya se guardan las fechas en el archivo. */
export function fechaAExcelSerial(fecha: Date): number {
  const epoca = Date.UTC(1899, 11, 30);
  const dia = Date.UTC(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  return Math.round((dia - epoca) / 86_400_000);
}

/** Extrae las filas de "Control de Facturas-Capex 25fEB" ya cargadas. */
export function extraerFacturas(wb: XLSX.WorkBook, nombreHoja: string): FacturaCapex[] {
  const filas = filasDeHoja(wb, nombreHoja);
  const facturas: FacturaCapex[] = [];
  for (let i = 1; i < filas.length; i++) {
    const fila = filas[i];
    const numeroFactura = aTexto(fila[COL_FACTURAS.numeroFactura]);
    const proyecto = aTexto(fila[COL_FACTURAS.proyecto]);
    if (!numeroFactura && !proyecto) continue;

    const fechaCruda = fila[COL_FACTURAS.periodoFacturado];
    let periodoTexto = "";
    let periodoISO = "";
    if (fechaCruda instanceof Date) {
      periodoTexto = fechaCruda.toLocaleDateString("es-PE", { year: "numeric", month: "2-digit", day: "2-digit" });
      periodoISO = fechaCruda.toISOString().slice(0, 10);
    } else {
      periodoTexto = aTexto(fechaCruda);
    }

    facturas.push({
      filaExcel: i + 1,
      periodoFacturado: periodoTexto,
      periodoFacturadoISO: periodoISO,
      recurso: aTexto(fila[COL_FACTURAS.recurso]),
      proveedor: aTexto(fila[COL_FACTURAS.proveedor]),
      responsable: aTexto(fila[COL_FACTURAS.responsable]),
      proyecto,
      monto: aNumero(fila[COL_FACTURAS.monto]),
      numeroFactura,
      registrado: aTexto(fila[COL_FACTURAS.registrado]),
      comentarios: aTexto(fila[COL_FACTURAS.comentarios]),
    });
  }
  return facturas;
}

/** Lee una celda cruda tal cual está (fórmula con "=" adelante, o su valor literal). */
export function leerCeldaCruda(wb: XLSX.WorkBook, nombreHoja: string, direccion: string): string | number {
  const hoja = wb.Sheets[nombreHoja];
  const celda = hoja?.[direccion];
  if (!celda) return "";
  if (celda.f) return `=${celda.f}`;
  return typeof celda.v === "number" ? celda.v : String(celda.v ?? "");
}

/** Extrae las filas de detalle de BD_CAPEX (lo que se está gastando/ejecutando actualmente). */
export function extraerProyectos(wb: XLSX.WorkBook, nombreHoja: string): ProyectoCapex[] {
  const filas = filasDeHoja(wb, nombreHoja);

  const proyectos: ProyectoCapex[] = [];
  for (let i = 1; i < filas.length; i++) {
    const fila = filas[i];
    const proyecto = aTexto(fila[COL_BD.proyecto]);
    const grupoNegocio = aTexto(fila[COL_BD.grupoNegocio]).toUpperCase();
    // Sin Grupo de Negocio no es un proyecto real: descarta filas vacías y también
    // subtotales sueltos que a veces quedan pegados dentro del rango de datos (ej. una
    // fila "TOTAL USD" sin grupo, con los montos del grupo anterior ya sumados).
    if (!grupoNegocio) continue;

    const real: number[] = [];
    const proyectado: number[] = [];
    for (let m = 0; m < 12; m++) {
      const colReal = COL_BD.primerMesReal + m * 2;
      const colProy = colReal + 1;
      real.push(aNumero(fila[colReal]));
      proyectado.push(aNumero(fila[colProy]));
    }

    proyectos.push({
      filaExcel: i + 1, // filas[i] es la fila (i+1) de la hoja (1-based, header en la fila 1)
      proyecto,
      subNegocio: aTexto(fila[COL_BD.subNegocio]),
      grupoNegocio,
      detalle: aTexto(fila[COL_BD.detalle]),
      avancePct: aTexto(fila[COL_BD.avancePct]),
      avance: aTexto(fila[COL_BD.avance]),
      categoria: aTexto(fila[COL_BD.categoria]),
      prioridad: aTexto(fila[COL_BD.prioridad]),
      status: aTexto(fila[COL_BD.status]),
      opex: aTexto(fila[COL_BD.opex]),
      recurso: aTexto(fila[COL_BD.recurso]),
      responsable: aTexto(fila[COL_BD.responsable]),
      tiempo: aTexto(fila[COL_BD.tiempo]),
      real,
      proyectado,
      presupuestoAprobado: aNumero(fila[COL_BD.presupuestoAprobado]),
    });
  }
  return proyectos;
}

/**
 * Extrae la proyección original de línea base ("Proyeccion 2026_vm"): el presupuesto
 * completo del año tal como se planificó al inicio, mes a mes.
 *
 * Los montos mensuales (columnas "Gasto Proyectado …") ya están en dólares — aunque la
 * celda los muestre con formato "S/", la propia hoja lo confirma: "Sub x Mes" (suma de
 * los 12 meses) es idéntico a "TOTAL USD" en cada fila, y la columna "TOTAL S/." se
 * calcula multiplicando ese total en USD por el tipo de cambio (celda A1), no al revés.
 * No hay que dividir nada.
 */
export function extraerProyeccionVM(wb: XLSX.WorkBook, nombreHoja: string): ItemConMeses[] {
  const filas = filasDeHoja(wb, nombreHoja);

  const items: ItemConMeses[] = [];
  for (let i = FILA_ENCABEZADO_VM + 1; i < filas.length; i++) {
    const fila = filas[i];
    const grupoNegocio = aTexto(fila[COL_VM.grupoNegocio]).toUpperCase();
    // Sin Grupo de Negocio no es una fila real: descarta encabezados/subtotales sueltos
    // que quedan pegados dentro del rango de datos (mismo criterio que en BD_CAPEX).
    if (!grupoNegocio) continue;

    const meses: number[] = [];
    for (let m = 0; m < 12; m++) {
      meses.push(aNumero(fila[COL_VM.primerMes + m]));
    }

    items.push({
      proyecto: aTexto(fila[COL_VM.proyecto]),
      detalle: aTexto(fila[COL_VM.detalle]),
      grupoNegocio,
      prioridad: aTexto(fila[COL_VM.prioridad]),
      meses,
    });
  }
  return items;
}
