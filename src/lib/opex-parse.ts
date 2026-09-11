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
  moneda: 15, // "PEN", "USD" o "EUR" — en qué moneda se ingresó originalmente el monto
  ruc: 16, // RUC del proveedor (solo aplica a proveedores peruanos) — opcional
  // 17 y 18: mismo criterio que montoSoles/tipoCambio, pero para facturas en Euros. Van
  // aparte (no las reemplazan) porque `tipoCambio` (columna 13) siempre guarda el tipo
  // de cambio Soles↔Dólar — se sigue usando tal cual para mostrar el equivalente en
  // Soles de CUALQUIER factura (incluida una en Euros); el tipo de cambio Euro↔Dólar es
  // un dato aparte, solo aplica a facturas ingresadas en Euros.
  montoEuros: 17, // Monto en Euros — solo cuando se ingresó en Euros
  tipoCambioEur: 18, // Tipo de cambio Euro→Dólar usado para convertir esta factura
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
  "RUC",
];

/** Encabezados agregados después de los anteriores (columnas 17-18) — para facturas en
 *  Euros. Se agregan por separado (mismo criterio: si la celda de esa columna en la
 *  fila 1 ya tiene algo, no se vuelve a escribir) porque este soporte se sumó en un
 *  cambio posterior. */
export const ENCABEZADOS_MONTO_EUROS_OPEX = ["Monto Euros", "Tipo de Cambio Euro"];

/** Serial de Excel (días desde 1899-12-30) → Date — el inverso de `fechaAExcelSerial` de
 *  capex-parse.ts. Hace falta porque las fechas que escribe la app (número plano, sin
 *  formato de celda aplicado) vuelven de Graph como número, no como Date — a diferencia
 *  de fechas que alguien tipeó directo en Excel con formato de fecha. */
function serialExcelAFecha(serial: number): Date {
  const epoca = Date.UTC(1899, 11, 30);
  return new Date(epoca + serial * 86_400_000);
}

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
  /** Equivalente en Soles sin IGV — si la factura se ingresó en Soles, es el valor que
   *  de verdad escribió la persona; si se ingresó en Dólares, es un cálculo de
   *  referencia (monto USD × tipoCambio DE ESA FILA, nunca el tipo de cambio actual) para
   *  que el reporte siempre muestre ambas monedas sin importar en cuál se registró.
   *  `null` solo en facturas de antes de que existiera este campo, sin tipoCambio
   *  guardado con qué calcularlo. */
  montoSoles: number | null;
  /** true si `montoSoles` es un cálculo de referencia (factura ingresada en USD),
   *  no el valor que la persona realmente escribió. */
  montoSolesEsCalculado: boolean;
  /** Tipo de cambio usado para convertir ESTA factura en particular — se guarda por
   *  fila (no solo el valor por defecto actual) para que el historial sea fiel incluso
   *  si el tipo de cambio por defecto cambia más adelante. */
  tipoCambio: number | null;
  /** Empresa de la línea de gasto — vacío en facturas registradas antes de este campo. */
  empresa: string;
  /** "PEN", "USD" o "EUR": en qué moneda ingresó la persona el monto — vacío en
   *  facturas registradas antes de que existiera el selector de moneda (esas siempre
   *  fueron en Soles, es la única moneda que aceptaba el formulario en ese momento). */
  moneda: "PEN" | "USD" | "EUR" | "";
  /** RUC del proveedor — vacío si no se conoce o si el proveedor no es peruano (el RUC
   *  es un identificador tributario solo de empresas registradas en Perú). */
  ruc: string;
  /** Monto en Euros tal como lo escribió la persona — solo cuando la factura se
   *  ingresó en Euros; `null` en cualquier otro caso. */
  montoEuros: number | null;
  /** Tipo de cambio Euro→Dólar usado para convertir ESTA factura en particular — solo
   *  cuando se ingresó en Euros; `null` en cualquier otro caso. */
  tipoCambioEur: number | null;
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
    let fecha: string;
    if (fechaCruda instanceof Date) {
      fecha = fechaCruda.toLocaleDateString("es-PE");
    } else if (typeof fechaCruda === "number" && fechaCruda > 0) {
      // Serial de Excel sin formato de fecha aplicado (así queda al escribirla desde la
      // app) — se convierte acá para mostrar una fecha real, no el número crudo.
      fecha = serialExcelAFecha(fechaCruda).toLocaleDateString("es-PE", { timeZone: "UTC" });
    } else {
      fecha = aTexto(fechaCruda);
    }

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

    // Si la factura se ingresó directo en Dólares, "Monto Soles" queda vacío en el Excel
    // (nunca se inventó un valor al registrarla). Para que el reporte y la pantalla
    // siempre muestren las dos monedas sin importar en cuál se registró, se calcula acá
    // — usando el tipo de cambio GUARDADO EN ESA MISMA FILA, nunca el actual: así, si el
    // tipo de cambio por defecto de la app cambia más adelante (hoy 3.4), las facturas
    // viejas siguen mostrando el valor correcto con el que de verdad se registraron.
    let montoSoles = montoSolesNum;
    let montoSolesEsCalculado = false;
    if (montoSoles == null && monto > 0) {
      montoSoles = Math.round(monto * (tipoCambioNum || TIPO_CAMBIO_POR_DEFECTO) * 100) / 100;
      montoSolesEsCalculado = true;
    }

    const montoEurosTxt = fila[COL_FACTURAS_OPEX.montoEuros];
    const tipoCambioEurTxt = fila[COL_FACTURAS_OPEX.tipoCambioEur];
    const montoEuros = montoEurosTxt !== "" && montoEurosTxt != null ? aNumero(montoEurosTxt) : null;
    const tipoCambioEur = tipoCambioEurTxt !== "" && tipoCambioEurTxt != null ? aNumero(tipoCambioEurTxt) : null;

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
      montoSoles,
      montoSolesEsCalculado,
      tipoCambio: tipoCambioNum,
      empresa: aTexto(fila[COL_FACTURAS_OPEX.empresa]),
      moneda: (aTexto(fila[COL_FACTURAS_OPEX.moneda]) as "PEN" | "USD" | "EUR" | ""),
      ruc: aTexto(fila[COL_FACTURAS_OPEX.ruc]),
      montoEuros,
      tipoCambioEur,
    });
  }
  return facturas;
}
