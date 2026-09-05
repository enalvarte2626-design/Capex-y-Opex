/** Una fila de BD_CAPEX ya tipada y con los montos convertidos a número (USD) — todas las
 *  columnas de la hoja, para el módulo de detalle. Los datos mensuales quedan crudos
 *  (real y proyectado por separado): el "mes de cierre" que decide dónde empieza el
 *  Forecast se aplica después, en `resolverProyecto`, para poder cambiarlo sin releer
 *  el Excel. */
export interface ProyectoCapex {
  /** Número de fila en la hoja de Excel (1-based), para poder editar esta fila puntual. */
  filaExcel: number;
  proyecto: string;
  subNegocio: string;
  grupoNegocio: string;
  detalle: string;
  avancePct: string;
  avance: string;
  categoria: string;
  prioridad: string;
  status: string;
  opex: string;
  recurso: string;
  responsable: string;
  tiempo: string;
  /** Gasto real por mes, en USD (índice 0 = enero … 11 = diciembre). */
  real: number[];
  /** Gasto proyectado por mes, en USD (índice 0 = enero … 11 = diciembre). */
  proyectado: number[];
  presupuestoAprobado: number;
}

/** ProyectoCapex ya resuelto para un mes de cierre concreto: real+proyectado combinados
 *  mes a mes, y Gasto real/Forecast/Diferencia calculados con ese corte. */
export interface ProyectoResuelto extends ProyectoCapex {
  meses: number[];
  gastoReal: number;
  forecast: number;
  diferencia: number;
}

/** Una fila de la hoja "Control de Facturas-Capex 25fEB". */
export interface FacturaCapex {
  filaExcel: number;
  periodoFacturado: string; // texto ya formateado (dd/mm/aaaa) para mostrar
  periodoFacturadoISO: string; // "aaaa-mm-dd", para poder editarlo con un <input type="date">
  recurso: string;
  proveedor: string;
  responsable: string;
  proyecto: string;
  monto: number;
  numeroFactura: string;
  registrado: string;
  comentarios: string;
}

export const NOMBRES_MES_CIERRE = [
  "Enero",
  "Febrero",
  "Marzo",
  "Abril",
  "Mayo",
  "Junio",
  "Julio",
  "Agosto",
  "Septiembre",
  "Octubre",
  "Noviembre",
  "Diciembre",
];

/**
 * Aplica el "mes de cierre" (1-12: último mes con Gasto Real cerrado) a un proyecto:
 * de ahí para atrás cuenta el Real, de ahí en adelante el Proyectado (Forecast) — misma
 * regla que ya usan las fórmulas de BD_CAPEX (GASTO REAL = suma de los 12 reales,
 * FORESCAST = suma del proyectado solo de los meses posteriores al cierre), pero
 * calculada aquí para poder mover el cierre sin tocar el Excel.
 */
export function resolverProyecto(p: ProyectoCapex, mesCierre: number): ProyectoResuelto {
  const meses = p.real.map((real, i) => real + (i >= mesCierre ? p.proyectado[i] : 0));
  const gastoReal = p.real.reduce((a, b) => a + b, 0);
  const forecast = p.proyectado.slice(mesCierre).reduce((a, b) => a + b, 0);
  const diferencia = p.presupuestoAprobado - gastoReal - forecast;
  return { ...p, meses, gastoReal, forecast, diferencia };
}

export function resolverProyectos(proyectos: ProyectoCapex[], mesCierre: number): ProyectoResuelto[] {
  return proyectos.map((p) => resolverProyecto(p, mesCierre));
}

/**
 * Encuentra a qué fila de BD_CAPEX (y qué mes) corresponde una factura, usando el mismo
 * texto que el módulo de Facturas ya escribe al registrarla: "Proyecto" = Detalle exacto
 * de esa fila, "Comentarios" = "Periodo {Mes}…". Si no hay una coincidencia única (ej.
 * facturas antiguas de antes de este módulo, con otro formato), devuelve null — esas
 * facturas no permiten editar el Monto porque no se puede ajustar el Gasto Real con
 * certeza.
 */
export function resolverFacturaABDCapex(
  factura: { proyecto: string; comentarios: string },
  proyectos: ProyectoCapex[]
): { filaProyecto: number; mes: number } | null {
  const textoProyecto = factura.proyecto.trim().toLowerCase();
  if (!textoProyecto) return null;

  const mesMatch = factura.comentarios.match(/Periodo\s+([A-Za-zÀ-ÿ]+)/i);
  if (!mesMatch) return null;
  const mesTexto = mesMatch[1].toLowerCase();
  const indiceMes = NOMBRES_MES_CIERRE.findIndex((n) => n.toLowerCase() === mesTexto);
  if (indiceMes === -1) return null;

  const coincidencias = proyectos.filter((p) => {
    const detalle = p.detalle.trim().toLowerCase();
    const nombre = p.proyecto.trim().toLowerCase();
    return (detalle && detalle === textoProyecto) || (!detalle && nombre === textoProyecto);
  });
  if (coincidencias.length !== 1) return null;

  return { filaProyecto: coincidencias[0].filaExcel, mes: indiceMes + 1 };
}

/** Lo mínimo que necesitan los cálculos de grupo/prioridad/trimestre — BD_CAPEX y la
 *  proyección de línea base (Proyeccion 2026_vm) comparten esta forma. */
export interface ItemConMeses {
  /** Nombre del proyecto, si la hoja lo trae (Proyeccion 2026_vm sí) — para poder
   *  identificar proyectos puntuales, ej. cuáles se corrieron de trimestre. */
  proyecto?: string;
  /** Detalle de la línea (Proyeccion 2026_vm trae varias filas por proyecto, una por
   *  Detalle/sub-tarea) — hace falta junto con `proyecto` para no mezclar líneas distintas
   *  del mismo proyecto al cruzarlas contra BD_CAPEX. */
  detalle?: string;
  grupoNegocio: string;
  prioridad: string;
  /** Gasto por mes en USD (índice 0 = enero … 11 = diciembre). */
  meses: number[];
}

export const NOMBRES_MES = [
  "Ene",
  "Feb",
  "Mar",
  "Abr",
  "May",
  "Jun",
  "Jul",
  "Ago",
  "Sep",
  "Oct",
  "Nov",
  "Dic",
];

export interface TotalesCapex {
  presupuestoAprobado: number;
  gastoReal: number;
  forecast: number;
  diferencia: number;
}

function sumarTotales(proyectos: ProyectoResuelto[]): TotalesCapex {
  return proyectos.reduce(
    (acc, p) => ({
      presupuestoAprobado: acc.presupuestoAprobado + p.presupuestoAprobado,
      gastoReal: acc.gastoReal + p.gastoReal,
      forecast: acc.forecast + p.forecast,
      diferencia: acc.diferencia + p.diferencia,
    }),
    { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 }
  );
}

export interface GastoPorGrupo extends TotalesCapex {
  grupoNegocio: string;
}

/** Gasto agrupado por Grupo de Negocio (Emisivo/Receptivo/Transversal), ya filtrado por prioridad. */
export function gastoPorGrupoNegocio(proyectos: ProyectoResuelto[]): GastoPorGrupo[] {
  const grupos = new Map<string, ProyectoResuelto[]>();
  for (const p of proyectos) {
    const clave = p.grupoNegocio || "SIN GRUPO";
    if (!grupos.has(clave)) grupos.set(clave, []);
    grupos.get(clave)!.push(p);
  }
  return Array.from(grupos.entries())
    .map(([grupoNegocio, filas]) => ({ grupoNegocio, ...sumarTotales(filas) }))
    .sort((a, b) => b.presupuestoAprobado - a.presupuestoAprobado);
}

/**
 * Regla fija de Avance: el % solo toma 4 valores posibles en BD_CAPEX (selección, no un
 * número libre) — igual que la columna "Avance" del Excel (=IF(E=0%,"No iniciado",...)).
 * Se centraliza aquí porque tanto el detalle de BD_CAPEX como el indicador del Dashboard
 * necesitan el mismo criterio para contar/filtrar por avance.
 */
/** Redondea cualquier valor libre al bucket fijo más cercano (0 / 0.3 / 0.8 / 1). */
export function valorAvance(avancePct: string): number | null {
  const valor = parseFloat(String(avancePct).replace(",", "."));
  if (Number.isNaN(valor)) return null;
  if (valor >= 1) return 1;
  if (valor >= 0.8) return 0.8;
  if (valor >= 0.3) return 0.3;
  return 0;
}

/** "Suspendido" no es un % de avance — se guarda en la columna Status (para no tocar el
 *  número de Avance ni romper fórmulas del Excel que puedan leer ese % en 0-1). */
export const MARCA_SUSPENDIDO = "Suspendido";

export function estaSuspendido(status: string): boolean {
  return status.trim().toLowerCase() === MARCA_SUSPENDIDO.toLowerCase();
}

/** Clave fija de avance de una fila: "0" / "0.3" / "0.8" / "1" / "suspendido" — o
 *  "sin_dato" cuando el % no cae en ningún bucket reconocido (vacío, texto libre, etc.). */
export function claveAvance(p: { avancePct: string; status: string }): string {
  if (estaSuspendido(p.status)) return "suspendido";
  const valor = valorAvance(p.avancePct);
  return valor === null ? "sin_dato" : String(valor);
}

export const ORDEN_AVANCE = ["0", "0.3", "0.8", "1", "suspendido", "sin_dato"] as const;

/** Mismas claves/orden que ORDEN_AVANCE — un solo lugar para el texto de cada bucket. */
export const ETIQUETAS_AVANCE: Record<string, string> = {
  "0": "0% · No iniciado",
  "0.3": "30% · Iniciado",
  "0.8": "80% · Por culminar",
  "1": "100% · Culminado",
  suspendido: "Suspendido",
  sin_dato: "Sin dato",
};

/** Mismo color por bucket en todas partes (badge de fila en BD_CAPEX, indicador del
 *  Dashboard) — así "Culminado" siempre se ve verde, "Suspendido" siempre rosa, etc. */
export const COLORES_AVANCE: Record<string, { bg: string; color: string }> = {
  "0": { bg: "var(--bg)", color: "var(--texto-suave)" },
  "0.3": { bg: "#e3edfa", color: "#0f6cbd" },
  "0.8": { bg: "#fdecdc", color: "var(--alerta)" },
  "1": { bg: "#e3f3e3", color: "var(--exito)" },
  suspendido: { bg: "#f3f2f1", color: "var(--peligro)" },
  sin_dato: { bg: "var(--bg)", color: "var(--texto-suave)" },
};

export interface AvancePorGrupo {
  grupoNegocio: string;
  /** Cantidad de líneas (Proyecto+Detalle) por clave de avance — ver claveAvance. */
  conteos: Record<string, number>;
  total: number;
}

/** Cuenta cuántas líneas de BD_CAPEX, por Grupo de Negocio, caen en cada estado de Avance
 *  (No iniciado/Iniciado/Por culminar/Culminado/Suspendido/Sin dato) — para el indicador
 *  del Dashboard. Cuenta líneas (Proyecto+Detalle), el mismo criterio que ya usa el resto
 *  del dashboard (ej. gastoPorGrupoNegocio), no proyectos únicos. */
export function avancePorGrupoNegocio<T extends { grupoNegocio: string; avancePct: string; status: string }>(
  items: T[]
): AvancePorGrupo[] {
  const grupos = new Map<string, Record<string, number>>();
  for (const p of items) {
    const grupo = p.grupoNegocio || "SIN GRUPO";
    if (!grupos.has(grupo)) grupos.set(grupo, {});
    const conteos = grupos.get(grupo)!;
    const clave = claveAvance(p);
    conteos[clave] = (conteos[clave] ?? 0) + 1;
  }
  return Array.from(grupos.entries())
    .map(([grupoNegocio, conteos]) => ({
      grupoNegocio,
      conteos,
      total: Object.values(conteos).reduce((a, b) => a + b, 0),
    }))
    .sort((a, b) => b.total - a.total);
}

export interface PrioridadDisponible {
  valor: string;
  cantidad: number;
}

/** Lista las prioridades presentes en los datos, para armar el filtro dinámicamente. */
export function prioridadesDisponibles<T extends { prioridad: string }>(items: T[]): PrioridadDisponible[] {
  const conteo = new Map<string, number>();
  for (const p of items) {
    const clave = p.prioridad || "Sin prioridad";
    conteo.set(clave, (conteo.get(clave) ?? 0) + 1);
  }
  return Array.from(conteo.entries())
    .map(([valor, cantidad]) => ({ valor, cantidad }))
    .sort((a, b) => a.valor.localeCompare(b.valor, "es", { numeric: true }));
}

export function filtrarPorPrioridad<T extends { prioridad: string }>(items: T[], prioridades: string[]): T[] {
  if (prioridades.length === 0) return items;
  const set = new Set(prioridades);
  return items.filter((p) => set.has(p.prioridad || "Sin prioridad"));
}

const TRIMESTRES = [
  [0, 1, 2],
  [3, 4, 5],
  [6, 7, 8],
  [9, 10, 11],
] as const;

export interface FilaTrimestral {
  prioridad: string;
  t: [number, number, number, number];
  total: number;
  /** Mismo gasto pero mes a mes (índice 0 = enero … 11 = diciembre) — para poder abrir
   *  el detalle de un trimestre que todavía mezcla meses reales con proyectados. */
  meses: number[];
}

export interface GrupoTrimestral {
  grupoNegocio: string;
  filas: FilaTrimestral[];
  subtotal: FilaTrimestral;
}

export interface PanoramaTrimestral {
  grupos: GrupoTrimestral[];
  total: FilaTrimestral;
}

export function mesesATrimestres(meses: number[]): [number, number, number, number] {
  return TRIMESTRES.map((idx) => idx.reduce<number>((suma, i) => suma + (meses[i] ?? 0), 0)) as [
    number,
    number,
    number,
    number,
  ];
}

function sumarFilas(filas: FilaTrimestral[]): FilaTrimestral {
  const t: [number, number, number, number] = [0, 0, 0, 0];
  const meses = Array(12).fill(0) as number[];
  let total = 0;
  for (const f of filas) {
    for (let i = 0; i < 4; i++) t[i] += f.t[i];
    f.meses.forEach((v, i) => {
      meses[i] += v;
    });
    total += f.total;
  }
  return { prioridad: "", t, total, meses };
}

/**
 * Panorama por Grupo de Negocio × Prioridad × Trimestre, igual a como se ve en la hoja
 * "RESUMEN CAPEX" (tablas "RESUMEN PANORAMA PROYECTADO/ACTUAL 2026"). Sirve tanto para la
 * proyección de línea base (Proyeccion 2026_vm) como para lo que se está gastando ahora
 * (BD_CAPEX) — ambas comparten la forma `ItemConMeses`.
 */
export function panoramaTrimestral(items: ItemConMeses[]): PanoramaTrimestral {
  const porGrupo = new Map<string, Map<string, ItemConMeses[]>>();
  for (const item of items) {
    const grupo = item.grupoNegocio || "SIN GRUPO";
    const prioridad = item.prioridad || "Sin prioridad";
    if (!porGrupo.has(grupo)) porGrupo.set(grupo, new Map());
    const porPrioridad = porGrupo.get(grupo)!;
    if (!porPrioridad.has(prioridad)) porPrioridad.set(prioridad, []);
    porPrioridad.get(prioridad)!.push(item);
  }

  const grupos: GrupoTrimestral[] = Array.from(porGrupo.entries())
    .map(([grupoNegocio, porPrioridad]) => {
      const filas: FilaTrimestral[] = Array.from(porPrioridad.entries())
        .map(([prioridad, filasItems]) => {
          const meses = Array(12).fill(0) as number[];
          for (const item of filasItems) {
            item.meses.forEach((v, i) => {
              meses[i] += v;
            });
          }
          const t = mesesATrimestres(meses);
          return { prioridad, t, total: t.reduce((a, b) => a + b, 0), meses };
        })
        .sort((a, b) => a.prioridad.localeCompare(b.prioridad, "es", { numeric: true }));
      return { grupoNegocio, filas, subtotal: { ...sumarFilas(filas), prioridad: "" } };
    })
    .sort((a, b) => b.subtotal.total - a.subtotal.total);

  const total = sumarFilas(grupos.map((g) => g.subtotal));
  return { grupos, total };
}
