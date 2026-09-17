import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { compararConReferencia } from "@/lib/compararCierre";
import { extraerPresupuestoOpex } from "@/lib/opex-parse";

export const dynamic = "force-dynamic";

/**
 * Igual que /api/capex/comparar-cierre, pero para el Presupuesto OPEX: ?itemId=...
 * (archivo, obligatorio), &nombre=... (para mostrar), &versionId=... (opcional — si no
 * viene, usa el contenido MÁS RECIENTE de ese archivo), &fechaVersion=... (opcional,
 * solo para mostrar), &cerrados=N (meses cerrados en ese punto de referencia) y
 * &cerradosActual=N (meses cerrados HOY, para el resumen por grupo).
 */
export async function GET(req: NextRequest) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const itemId = req.nextUrl.searchParams.get("itemId")?.trim();
  const nombre = req.nextUrl.searchParams.get("nombre")?.trim();
  const versionId = req.nextUrl.searchParams.get("versionId")?.trim() || undefined;
  const fechaVersion = req.nextUrl.searchParams.get("fechaVersion")?.trim() || undefined;
  const cerrados = Number(req.nextUrl.searchParams.get("cerrados") ?? "0");
  const cerradosActual = Number(req.nextUrl.searchParams.get("cerradosActual") ?? "0");

  if (!itemId || !nombre) {
    return NextResponse.json({ error: "Falta indicar contra qué archivo comparar." }, { status: 400 });
  }
  if (!Number.isInteger(cerrados) || cerrados < 0 || cerrados > 12) {
    return NextResponse.json({ error: "Meses cerrados en la referencia inválido (debe ser 0-12)." }, { status: 400 });
  }
  if (!Number.isInteger(cerradosActual) || cerradosActual < 0 || cerradosActual > 12) {
    return NextResponse.json({ error: "Meses cerrados ahora inválido (debe ser 0-12)." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const resultado = await compararConReferencia(
      config,
      archivo,
      hoja,
      { driveId: archivo.driveId, itemId, nombre, versionId, fechaVersion },
      cerrados,
      cerradosActual,
      undefined,
      extraerPresupuestoOpex
    );
    return NextResponse.json(resultado);
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
