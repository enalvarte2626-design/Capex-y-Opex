/**
 * Única fuente de verdad de qué campos de BD_CAPEX se pueden editar desde la app y a
 * qué columna de Excel corresponde cada uno. Compartido entre el frontend (qué mostrar
 * como editable) y la ruta API que escribe la celda (qué se permite escribir).
 */

/** Convierte un índice de columna 0-based (0=A, 1=B, 25=Z, 26=AA…) a letra de Excel. */
export function columnaALetra(indice: number): string {
  let n = indice + 1;
  let letra = "";
  while (n > 0) {
    const resto = (n - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    n = Math.floor((n - 1) / 26);
  }
  return letra;
}

// Mismos índices que COL_BD en capex-parse.ts — si esas columnas se mueven en el Excel,
// hay que actualizar ambos lados.
const COL = {
  proyecto: 0,
  detalle: 3,
  prioridad: 7,
  status: 8,
  responsable: 11,
  avancePct: 4,
  presupuestoAprobado: 37,
  primerMesReal: 13,
} as const;

export type TipoCampo = "texto" | "numero" | "avance";

export interface DefinicionCampo {
  columna: string; // letra de Excel
  tipo: TipoCampo;
  /** Longitud máxima para campos de texto (evita pegar bloques enormes por error). */
  largoMaximo?: number;
}

// "presupuestoAprobado" a propósito NO está aquí: se puede fijar al crear un proyecto
// nuevo (agregar-proyecto/route.ts, su propio flujo), pero no se permite editar después
// desde Detalle BD_CAPEX — es un valor aprobado, no algo para ajustar sobre la marcha.
const CAMPOS_SIMPLES: Record<string, DefinicionCampo> = {
  proyecto: { columna: columnaALetra(COL.proyecto), tipo: "texto", largoMaximo: 200 },
  detalle: { columna: columnaALetra(COL.detalle), tipo: "texto", largoMaximo: 500 },
  status: { columna: columnaALetra(COL.status), tipo: "texto", largoMaximo: 100 },
  prioridad: { columna: columnaALetra(COL.prioridad), tipo: "texto", largoMaximo: 10 },
  responsable: { columna: columnaALetra(COL.responsable), tipo: "texto", largoMaximo: 100 },
  avancePct: { columna: columnaALetra(COL.avancePct), tipo: "avance" },
};

const VALORES_AVANCE_VALIDOS = [0, 0.3, 0.8, 1];

/**
 * Resuelve un identificador de campo a su columna de Excel + tipo esperado. Los meses
 * usan el identificador "real:0".."real:11" / "proyectado:0".."proyectado:11" (0=enero).
 */
export function resolverCampo(campo: string): DefinicionCampo | null {
  if (CAMPOS_SIMPLES[campo]) return CAMPOS_SIMPLES[campo];

  const mesMatch = campo.match(/^(real|proyectado):(\d+)$/);
  if (mesMatch) {
    const tipo = mesMatch[1] as "real" | "proyectado";
    const mes = Number(mesMatch[2]);
    if (mes < 0 || mes > 11) return null;
    const colReal = COL.primerMesReal + mes * 2;
    const col = tipo === "real" ? colReal : colReal + 1;
    return { columna: columnaALetra(col), tipo: "numero" };
  }
  return null;
}

/** Valida y normaliza el valor recibido según el tipo del campo. Devuelve null si no es válido. */
export function validarValor(def: DefinicionCampo, valorCrudo: unknown): number | string | null {
  if (def.tipo === "avance") {
    const n = Number(valorCrudo);
    return VALORES_AVANCE_VALIDOS.includes(n) ? n : null;
  }
  if (def.tipo === "numero") {
    const n = Number(valorCrudo);
    return Number.isFinite(n) ? n : null;
  }
  // texto
  const texto = String(valorCrudo ?? "").trim();
  if (def.largoMaximo && texto.length > def.largoMaximo) return null;
  return texto;
}
