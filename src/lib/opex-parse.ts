import * as XLSX from "xlsx";
import type { ProyectoCapex } from "./capex";
import { TIPO_CAMBIO_POR_DEFECTO } from "./opex-constantes";

/**
 * Índices de columna (0-based) de "Presupuesto 2026" — el equivalente a BD_CAPEX para
 * OPEX, con la diferencia de que es una Tabla de Excel real (fórmulas con
 * `Tabla1[...]`) en vez de rangos sueltos. Encabezado en la fila 1, datos desde la 2.
 */
export const COL_PPTO_OPEX = {
  empresa: 2, // C
  grupoGasto: 3, // D
  subgrupoGasto: 4, // E
  lineaGasto: 5, // F: "Lineas Opex-2025" — la línea de gasto específica
  status: 6, // G
  moneda: 7, // H
  detalle: 8, // I
  responsable: 9, // J
  primerMesProyectado: 10, // K: Gasto Proyectado Enero (luego alternan Proy/Real x 12 meses)
  primerMesReal: 11, // L: Gasto Real Enero
  presupuestoAprobado: 34, // AI: PPTO APROBADO 2026
} as const;

/**
 * Índices de columna (0-based) de la hoja nueva "Facturas Opex - App".
 *
 * `montoSoles` y `tipoCambio` se agregaron para la regularización de facturas en Soles
 * sin IGV: `monto` (columna F, la que de verdad mueve el Gasto Real) siempre queda en
 * USD, calculado como `montoSoles / tipoCambio` — nunca se pide el USD directo. `empresa`
 * se agregó después, para que la base de gastos quede completa con todos los campos que
 * ya se piden en el formulario (antes solo se usaba para filtrar, no se guardaba). Todas
 * van al final de la fila para no correr ninguna columna existente: las facturas
 * registradas antes de cada cambio simplemente quedan con esas celdas vacías.
 */
export const COL_FACTURAS_OPEX = {
  fecha: 0,
  grupoGasto: 1,
  subgrupoGasto: 2,
  lineaGasto: 3,
  filaPresupuesto: 4, // referencia directa a la fila de "Presupuesto 2026" — sin ambigüedad
  mes: 5, // 1-12
  monto: 6, // USD — siempre calculado, nunca ingresado directo
  proveedor: 7,
  numeroComprobante: 8,
  responsable: 9,
  comentario: 10,
  registrado: 11,
  montoSoles: 12, // Monto en Soles sin IGV — solo cuando la factura se ingresó en Soles
  tipoCambio: 13, // Tipo de cambio usado para convertir esta factura en particular
  empresa: 14, // Empresa de la línea de gasto elegida (viene de Presupuesto 2026)
  moneda: 15, // "PEN" o "USD" — en qué moneda se ingresó originalmente el monto
} as const;

export const ENCABEZADOS_FACTURAS_OPEX = [
  "Fecha",
  "Grupo de Gasto",
  "Subgrupo de Gasto",
  "Línea de Gasto",
  "Fila Presupuesto",
  "Mes",
  "Monto (USD)",
  "Proveedor",
  "N° Comprobante",
  "Responsable",
  "Comentario",
  "Registrado",
  "Monto Soles (sin IGV)",
  "Tipo de Cambio",
  "Empresa",
  "Moneda ingresada",
];

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

function filasDeHoja(wb: XLSX.WorkBook, nombreHoja: string): unknown[][] {
  const hoja = wb.Sheets[nombreHoja];
  if (!hoja) {
    const disponibles = wb.SheetNames.join(", ");
    throw new Error(`No se encontró la hoja "${nombreHoja}" en el archivo. Hojas disponibles: ${disponibles}.`);
  }
  return XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true, defval: "" });
}

/**
 * Extrae "Presupuesto 2026" reusando el mismo molde `ProyectoCapex` que ya usa CAPEX
 * (mismos cálculos de real+proyectado por mes, presupuesto/gasto real/forecast) — solo
 * cambia qué significa cada campo:
 * `proyecto` = Línea de Gasto, `grupoNegocio` = Grupo de Gasto, `prioridad` = Subgrupo de
 * Gasto, `subNegocio` = Empresa, `detalle` = Detalle libre, `status`/`responsable` tal cual.
 */
export function extraerPresupuestoOpex(wb: XLSX.WorkBook, nombreHoja: string): ProyectoCapex[] {
  const filas = filasDeHoja(wb, nombreHoja);
  const lineas: ProyectoCapex[] = [];

  for (let i = 1; i < filas.length; i++) {
    const fila = filas[i];
    const grupoGasto = aTexto(fila[COL_PPTO_OPEX.grupoGasto]).toUpperCase();
    // Sin Grupo de Gasto no es una línea real: descarta filas vacías al final de la tabla.
    if (!grupoGasto) continue;

    const real: number[] = [];
    const proyectado: number[] = [];
    for (let m = 0; m < 12; m++) {
      const colProy = COL_PPTO_OPEX.primerMesProyectado + m * 2;
      const colReal = COL_PPTO_OPEX.primerMesReal + m * 2;
      proyectado.push(aNumero(fila[colProy]));
      real.push(aNumero(fila[colReal]));
    }

    lineas.push({
      filaExcel: i + 1,
      proyecto: aTexto(fila[COL_PPTO_OPEX.lineaGasto]),
      subNegocio: aTexto(fila[COL_PPTO_OPEX.empresa]),
      grupoNegocio: grupoGasto,
      detalle: aTexto(fila[COL_PPTO_OPEX.detalle]),
      avancePct: "0",
      avance: "",
      categoria: "",
      prioridad: aTexto(fila[COL_PPTO_OPEX.subgrupoGasto]),
      status: aTexto(fila[COL_PPTO_OPEX.status]),
      opex: "",
      recurso: "",
      responsable: aTexto(fila[COL_PPTO_OPEX.responsable]),
      tiempo: "",
      real,
      proyectado,
      presupuestoAprobado: aNumero(fila[COL_PPTO_OPEX.presupuestoAprobado]),
    });
  }
  return lineas;
}

export interface FacturaOpex {
  filaExcel: number;
  fecha: string;
  grupoGasto: string;
  subgrupoGasto: string;
  lineaGasto: string;
  filaPresupuesto: number | null;
  mes: number | null;
  monto: number;
  proveedor: string;
  numeroComprobante: string;
  responsable: string;
  comentario: string;
  registrado: string;
  /** Monto original en Soles sin IGV que ingresó la persona — null en facturas
   *  registradas antes de este campo (esas solo tienen `monto` en USD). */
  montoSoles: number | null;
  /** Tipo de cambio usado para convertir ESTA factura en particular — se guarda por
   *  fila (no solo el valor por defecto actual) para que el historial sea fiel incluso
   *  si el tipo de cambio por defecto cambia más adelante. */
  tipoCambio: number | null;
  /** Empresa de la línea de gasto — vacío en facturas registradas antes de este campo. */
  empresa: string;
  /** "PEN" o "USD": en qué moneda ingresó la persona el monto — vacío en facturas
   *  registradas antes de que existiera el selector de moneda (esas siempre fueron en
   *  Soles, es la única moneda que aceptaba el formulario en ese momento). */
  moneda: "PEN" | "USD" | "";
}

/** Extrae "Facturas Opex - App" — si la hoja todavía no existe (nadie ha registrado
 *  ninguna factura desde la app todavía), devuelve una lista vacía. */
export function extraerFacturasOpex(wb: XLSX.WorkBook, nombreHoja: string): FacturaOpex[] {
  const hoja = wb.Sheets[nombreHoja];
  if (!hoja) return [];
  const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true, defval: "" }) as unknown[][];
  const facturas: FacturaOpex[] = [];

  for (let i = 1; i < filas.length; i++) {
    const fila = filas[i];
    const proveedor = aTexto(fila[COL_FACTURAS_OPEX.proveedor]);
    const lineaGasto = aTexto(fila[COL_FACTURAS_OPEX.lineaGasto]);
    if (!proveedor && !lineaGasto) continue;

    const fechaCruda = fila[COL_FACTURAS_OPEX.fecha];
    const fecha = fechaCruda instanceof Date ? fechaCruda.toLocaleDateString("es-PE") : aTexto(fechaCruda);

    const filaPresupuestoTxt = fila[COL_FACTURAS_OPEX.filaPresupuesto];
    const mesTxt = fila[COL_FACTURAS_OPEX.mes];
    const montoSolesTxt = fila[COL_FACTURAS_OPEX.montoSoles];
    const tipoCambioTxt = fila[COL_FACTURAS_OPEX.tipoCambio];
    const montoSolesNum = montoSolesTxt !== "" && montoSolesTxt != null ? aNumero(montoSolesTxt) : null;
    const tipoCambioNum = tipoCambioTxt !== "" && tipoCambioTxt != null ? aNumero(tipoCambioTxt) : null;

    // Respaldo para filas con Monto (USD) en 0 pero con Soles sin IGV sí guardado — pasó
    // con las primeras facturas registradas mientras un bug dejaba esa celda en blanco
    // (ver comentario en opex-constantes.ts). Nunca se sobrescribe el Excel con esto:
    // es solo para que la persona vea el dólar correcto en pantalla mientras corrige o
    // vuelve a guardar esa fila.
    let monto = aNumero(fila[COL_FACTURAS_OPEX.monto]);
    if (monto === 0 && montoSolesNum && montoSolesNum > 0) {
      monto = Math.round((montoSolesNum / (tipoCambioNum || TIPO_CAMBIO_POR_DEFECTO)) * 100) / 100;
    }

    facturas.push({
      filaExcel: i + 1,
      fecha,
      grupoGasto: aTexto(fila[COL_FACTURAS_OPEX.grupoGasto]),
      subgrupoGasto: aTexto(fila[COL_FACTURAS_OPEX.subgrupoGasto]),
      lineaGasto,
      filaPresupuesto: filaPresupuestoTxt ? Number(filaPresupuestoTxt) : null,
      mes: mesTxt ? Number(mesTxt) : null,
      monto,
      proveedor,
      numeroComprobante: aTexto(fila[COL_FACTURAS_OPEX.numeroComprobante]),
      responsable: aTexto(fila[COL_FACTURAS_OPEX.responsable]),
      comentario: aTexto(fila[COL_FACTURAS_OPEX.comentario]),
      registrado: aTexto(fila[COL_FACTURAS_OPEX.registrado]),
      montoSoles: montoSolesNum,
      tipoCambio: tipoCambioNum,
      empresa: aTexto(fila[COL_FACTURAS_OPEX.empresa]),
      moneda: (aTexto(fila[COL_FACTURAS_OPEX.moneda]) as "PEN" | "USD" | ""),
    });
  }
  return facturas;
}
