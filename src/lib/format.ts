const formatoMoneda = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const formatoMoneda2Dec = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export function moneda(valor: number): string {
  return `$${formatoMoneda.format(Math.round(valor))}`;
}

/** Con 2 decimales, para montos que se editan (ej. gasto real/proyectado por mes). */
export function moneda2(valor: number): string {
  return `$${formatoMoneda2Dec.format(valor)}`;
}

/** Formato abreviado en miles, igual que la fila "META" del Excel: "$315K"… */
export function monedaK(valor: number): string {
  const signo = valor < 0 ? "-" : "";
  return `${signo}$${Math.round(Math.abs(valor) / 1000)}K`;
}

/** Equivalente en Soles de un monto en USD — solo de referencia visual, nunca se guarda ni
 *  se usa para calcular nada; el USD sigue siendo la fuente de verdad del Excel. */
export function soles(valorUsd: number, tipoCambio: number): string {
  return `S/ ${(valorUsd * tipoCambio).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

/** Igual que monedaK pero en Soles — para los ejes/tooltips de gráficos cuando "Habilitar
 *  en Soles" está activo. */
export function solesK(valorUsd: number, tipoCambio: number): string {
  const valor = valorUsd * tipoCambio;
  const signo = valor < 0 ? "-" : "";
  return `${signo}S/ ${Math.round(Math.abs(valor) / 1000)}K`;
}
