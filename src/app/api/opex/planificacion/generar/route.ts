import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { generarBorradorOpex } from "@/lib/planificacionOpex";

export const dynamic = "force-dynamic";

/** Genera (o regenera desde cero) el borrador de planificación de OPEX, copiando la
 *  lista de líneas vigente hoy en Presupuesto 2026, en blanco. */
export async function POST() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const resultado = await generarBorradorOpex(config, archivo, hoja);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
