import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, escribirCelda, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { resolverCampo, validarValor } from "@/lib/capex-editable";
import { HOJA_PLANIFICACION_CAPEX } from "@/lib/planificacionCapex";

export const dynamic = "force-dynamic";

/** Igual que /api/capex/celda, pero escribe en la hoja de planificación (borrador) en
 *  vez de BD_CAPEX — mismos campos editables (capex-editable.ts), incluido
 *  "presupuestoAprobado" que en BD_CAPEX no se puede tocar después de crear el
 *  proyecto: en el borrador sí, porque todavía no es un presupuesto aprobado de verdad. */
export async function POST(request: Request) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = await request.json().catch(() => null);
  const fila = Number(cuerpo?.fila);
  const campo = String(cuerpo?.campo ?? "");

  if (!Number.isInteger(fila) || fila < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }

  // "presupuestoAprobado" a propósito no está en CAMPOS_SIMPLES de capex-editable.ts
  // (no se permite tocar en BD_CAPEX) — acá sí se acepta, resolviendo su columna a mano
  // con el mismo criterio que usa /api/capex/agregar-proyecto (columna AL).
  const definicion = campo === "presupuestoAprobado" ? { columna: "AL", tipo: "numero" as const } : resolverCampo(campo);
  if (!definicion) {
    return NextResponse.json({ error: `Campo "${campo}" no es editable.` }, { status: 400 });
  }

  const valor = validarValor(definicion, cuerpo?.valor);
  if (valor === null) {
    return NextResponse.json({ error: "Valor inválido para este campo." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    await escribirCelda(config, archivo, HOJA_PLANIFICACION_CAPEX, `${definicion.columna}${fila}`, valor);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
