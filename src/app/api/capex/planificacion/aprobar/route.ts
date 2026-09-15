import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { aprobarYActivarCapex } from "@/lib/planificacionCapex";
import { anioActivoCapex } from "@/lib/anio";

export const dynamic = "force-dynamic";

/**
 * "Aprobar y activar" CAPEX: el borrador de planificación pasa a ser BD_CAPEX, y el
 * BD_CAPEX saliente queda archivado con el año en el nombre. Ver
 * planificacionCapex.ts#aprobarYActivarCapex — todo por renombre de hoja, nunca
 * copiando ni borrando filas.
 */
export async function POST() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const anioActivo = await anioActivoCapex();
    const resultado = await aprobarYActivarCapex(config, archivo, hoja, anioActivo);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
