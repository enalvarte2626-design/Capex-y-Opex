/**
 * Detecta problemas reales en los datos del Excel (BD_CAPEX / Presupuesto 2026) que hoy
 * quedan invisibles: una celda con un error de fórmula (#REF!, #VALUE!, etc.) se convierte
 * en $0 silenciosamente al leerla, y una fórmula que suma de más (ej. una columna entera
 * en vez de solo esa fila) da un número gigante pero perfectamente numérico — nada en la
 * app avisaba de ninguno de los dos casos. Esto es justo lo que pasó en el Excel real de
 * OPEX: una fila con un #REF! en el Gasto Real de un mes, y otra con el Forecast inflado
 * a más de $1,000,000 por una fórmula mal copiada.
 */

const NOMBRES_MES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

/** Los valores de error que Excel escribe tal cual en la celda cuando una fórmula falla. */
export function esErrorExcel(valor: unknown): boolean {
  if (typeof valor !== "string") return false;
  return /^#(REF|VALUE|DIV\/0|N\/A|NAME\?|NULL|NUM)!?$/i.test(valor.trim());
}

interface DatosParaAdvertencias {
  /** Valores crudos (antes de convertir a número) — hace falta el texto original para
   *  poder detectar "#REF!" y compañía, que `aNumero()` ya convirtió a 0. */
  realCrudo: unknown[];
  proyectadoCrudo: unknown[];
  presupuestoAprobadoCrudo: unknown;
  /** Los mismos datos ya convertidos a número, para la heurística de desproporción. */
  real: number[];
  proyectado: number[];
  presupuestoAprobado: number;
}

/**
 * Arma la lista de advertencias de una línea presupuestal. Vacía si no hay nada raro.
 * Dos tipos de aviso:
 * 1) Una celda con error de fórmula (#REF!, #VALUE!...) — se está contando como $0 sin que
 *    nadie se entere.
 * 2) Un solo mes cuyo Gasto Real o Proyectado, por sí solo, ya supera varias veces el
 *    Presupuesto Aprobado del año completo — típico de una fórmula que suma de más. Solo
 *    aplica cuando SÍ hay presupuesto aprobado (>0): una línea en $0 de presupuesto con
 *    algo de gasto real es normal (proyecto nuevo aún sin presupuesto formal asignado), no
 *    un error — de ahí que este mismo chequeo sin ese resguardo diera falsos positivos en
 *    el Excel.
 */
export function detectarAdvertencias(d: DatosParaAdvertencias): string[] {
  const advertencias: string[] = [];

  if (esErrorExcel(d.presupuestoAprobadoCrudo)) {
    advertencias.push(
      `Presupuesto Aprobado tiene un error en el Excel (${d.presupuestoAprobadoCrudo}) — se está contando como $0.`
    );
  }

  for (let m = 0; m < 12; m++) {
    if (esErrorExcel(d.realCrudo[m])) {
      advertencias.push(
        `Gasto Real de ${NOMBRES_MES[m]} tiene un error en el Excel (${d.realCrudo[m]}) — se está contando como $0.`
      );
    }
    if (esErrorExcel(d.proyectadoCrudo[m])) {
      advertencias.push(
        `Gasto Proyectado de ${NOMBRES_MES[m]} tiene un error en el Excel (${d.proyectadoCrudo[m]}) — se está contando como $0.`
      );
    }
  }

  if (d.presupuestoAprobado > 0) {
    const revisa = (valor: number, mes: number, etiqueta: string) => {
      // Umbral doble (ratio + monto absoluto) para no marcar líneas con presupuesto chico
      // donde cualquier variación normal ya se ve como "varias veces" el presupuesto.
      if (valor > d.presupuestoAprobado * 3 && valor - d.presupuestoAprobado > 1000) {
        advertencias.push(
          `${etiqueta} de ${NOMBRES_MES[mes]} ($${valor.toLocaleString("es-PE", { maximumFractionDigits: 0 })}) parece ` +
            `desproporcionado frente al Presupuesto Aprobado ($${d.presupuestoAprobado.toLocaleString("es-PE", { maximumFractionDigits: 0 })}) — revisar la fórmula en el Excel.`
        );
      }
    };
    d.real.forEach((v, m) => revisa(v, m, "Gasto Real"));
    d.proyectado.forEach((v, m) => revisa(v, m, "Gasto Proyectado"));
  }

  return advertencias;
}
