import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { agregarLineaBorradorOpex } from "@/lib/planificacionOpex";

export const dynamic = "force-dynamic";

/** Agrega una línea nueva al borrador de planificación de OPEX — mismos campos que pide
 *  "+ Agregar línea" en Presupuesto OPEX, escribiendo en la hoja de planificación. */
export async function POST(req: NextRequest) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const body = await req.json().catch(() => null);
  const empresa = String(body?.empresa ?? "").trim();
  const grupoGasto = String(body?.grupoGasto ?? "").trim();
  const subgrupoGasto = String(body?.subgrupoGasto ?? "").trim();
  const lineaGasto = String(body?.lineaGasto ?? "").trim();
  const moneda = String(body?.moneda ?? "USD").trim() || "USD";
  const detalle = String(body?.detalle ?? "").trim();
  const responsable = String(body?.responsable ?? "").trim();
  const presupuestoAprobado = Number(body?.presupuestoAprobado ?? 0);

  if (!empresa) return NextResponse.json({ error: "Falta la Empresa." }, { status: 400 });
  if (!grupoGasto) return NextResponse.json({ error: "Falta el Grupo de Gasto." }, { status: 400 });
  if (!subgrupoGasto) return NextResponse.json({ error: "Falta el Subgrupo de Gasto." }, { status: 400 });
  if (!lineaGasto) return NextResponse.json({ error: "Falta la Línea de Gasto." }, { status: 400 });
  if (!Number.isFinite(presupuestoAprobado) || presupuestoAprobado < 0) {
    return NextResponse.json({ error: "El Presupuesto Aprobado debe ser un número válido." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const resultado = await agregarLineaBorradorOpex(config, archivo, {
      empresa,
      grupoGasto,
      subgrupoGasto,
      lineaGasto,
      moneda,
      detalle,
      responsable,
      presupuestoAprobado,
    });
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
