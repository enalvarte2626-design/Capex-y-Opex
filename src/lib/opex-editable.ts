/**
 * Registro de campos editables de "Presupuesto 2026" (OPEX) — mismo patrón que
 * capex-editable.ts, reusando sus tipos y su conversión índice→letra.
 */
import { columnaALetra, validarValor, type DefinicionCampo, type TipoCampo } from "./capex-editable";
import { COL_PPTO_OPEX } from "./opex-parse";

export type { TipoCampo, DefinicionCampo };
export { columnaALetra, validarValor };

const CAMPOS_SIMPLES: Record<string, DefinicionCampo> = {
  detalle: { columna: columnaALetra(COL_PPTO_OPEX.detalle), tipo: "texto", largoMaximo: 500 },
  status: { columna: columnaALetra(COL_PPTO_OPEX.status), tipo: "texto", largoMaximo: 100 },
  responsable: { columna: columnaALetra(COL_PPTO_OPEX.responsable), tipo: "texto", largoMaximo: 100 },
  presupuestoAprobado: { columna: columnaALetra(COL_PPTO_OPEX.presupuestoAprobado), tipo: "numero" },
};

/**
 * Resuelve un identificador de campo a su columna de Excel + tipo esperado. Los meses
 * usan "real:0".."real:11" / "proyectado:0".."proyectado:11" (0=enero).
 */
export function resolverCampoOpex(campo: string): DefinicionCampo | null {
  if (CAMPOS_SIMPLES[campo]) return CAMPOS_SIMPLES[campo];

  const mesMatch = campo.match(/^(real|proyectado):(\d+)$/);
  if (mesMatch) {
    const tipo = mesMatch[1] as "real" | "proyectado";
    const mes = Number(mesMatch[2]);
    if (mes < 0 || mes > 11) return null;
    const col =
      tipo === "real" ? COL_PPTO_OPEX.primerMesReal + mes * 2 : COL_PPTO_OPEX.primerMesProyectado + mes * 2;
    return { columna: columnaALetra(col), tipo: "numero" };
  }
  return null;
}
