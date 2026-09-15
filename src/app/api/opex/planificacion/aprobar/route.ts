import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { aprobarYActivarOpex } from "@/lib/planificacionOpex";
import { anioActivoOpex } from "@/lib/anio";

export const dynamic = "force-dynamic";

/** "Aprobar y activar" OPEX: el borrador pasa a ser Presupuesto 2026 (o el nombre que
 *  tenga hoy la hoja en vivo), y la hoja saliente queda archivada con el año en el
 *  nombre. Ver planificacionOpex.ts#aprobarYActivarOpex. */
export async function POST() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const anioActivo = await anioActivoOpex();
    const resultado = await aprobarYActivarOpex(config, archivo, hoja, anioActivo);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
