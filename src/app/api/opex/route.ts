import { NextResponse } from "next/server";
import { ErrorSharePoint, descargarContenido, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { camposFaltantesOpex } from "@/lib/sharepoint";
import { extraerPresupuestoOpex } from "@/lib/opex-parse";
import { leerWorkbook } from "@/lib/capex-parse";

export const dynamic = "force-dynamic";

/** Igual que /api/capex: siempre lee el Excel en vivo, sin caché ni base de datos propia. */
export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const lineas = extraerPresupuestoOpex(wb, hoja);

    return NextResponse.json({
      lineas,
      archivo: archivo.nombre,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}