import { soles } from "@/lib/format";

/**
 * Línea gris chica con el equivalente en Soles de un monto en USD — se usa debajo de
 * cualquier importe en dólares cuando "Habilitar en Soles" está activo, mismo criterio en
 * todas las tarjetas/tablas/gráficos de CAPEX y OPEX. Solo de referencia visual: nunca
 * afecta ningún cálculo ni se guarda en el Excel (la fuente de verdad sigue siendo USD).
 */
export default function MontoSoles({
  valorUsd,
  tipoCambio,
  mostrarSoles,
  className = "block text-xs font-normal",
}: {
  valorUsd: number;
  tipoCambio: number;
  mostrarSoles: boolean;
  className?: string;
}) {
  if (!mostrarSoles) return null;
  return (
    <span className={className} style={{ color: "var(--texto-suave)" }}>
      {soles(valorUsd, tipoCambio)}
    </span>
  );
}
