import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { leerBorradorOpex } from "@/lib/planificacionOpex";
import { anioActivoOpex } from "@/lib/anio";

export const dynamic = "force-dynamic";

/** Lee la hoja de planificación (borrador) de OPEX — vacía si nadie la generó todavía
 *  este año. */
export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const [lineas, anioActivo] = await Promise.all([leerBorradorOpex(config, archivo), anioActivoOpex()]);
    return NextResponse.json({
      lineas,
      anioActivo,
      anioBorrador: anioActivo + 1,
      archivo: archivo.nombre,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
