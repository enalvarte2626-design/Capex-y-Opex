import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { agregarLineaBorradorCapex } from "@/lib/planificacionCapex";

export const dynamic = "force-dynamic";

/** Agrega un proyecto nuevo al borrador de planificación de CAPEX — mismos campos que
 *  /api/capex/agregar-proyecto, pero sin Gasto Real/Proyectado inicial (todavía no
 *  arrancó el año) y escribiendo en la hoja de planificación, no en BD_CAPEX. */
export async function POST(req: NextRequest) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const body = await req.json().catch(() => null);
  const proyecto = String(body?.proyecto ?? "").trim();
  const subNegocio = String(body?.subNegocio ?? "").trim();
  const grupoNegocio = String(body?.grupoNegocio ?? "").trim().toUpperCase();
  const detalle = String(body?.detalle ?? "").trim();
  const prioridad = String(body?.prioridad ?? "").trim();
  const presupuestoAprobado = Number(body?.presupuestoAprobado ?? 0);

  if (!["EMISIVO", "RECEPTIVO", "TRANSVERSAL"].includes(grupoNegocio)) {
    return NextResponse.json({ error: "Grupo de Negocio debe ser Emisivo, Receptivo o Transversal." }, { status: 400 });
  }
  if (!prioridad) return NextResponse.json({ error: "Falta la Prioridad." }, { status: 400 });
  if (!Number.isFinite(presupuestoAprobado) || presupuestoAprobado < 0) {
    return NextResponse.json({ error: "El Presupuesto Aprobado debe ser un número válido." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const resultado = await agregarLineaBorradorCapex(config, archivo, {
      proyecto,
      subNegocio,
      grupoNegocio,
      detalle,
      prioridad,
      presupuestoAprobado,
    });
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
