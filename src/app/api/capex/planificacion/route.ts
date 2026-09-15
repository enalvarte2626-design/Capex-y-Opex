import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { leerBorradorCapex } from "@/lib/planificacionCapex";
import { anioActivoCapex } from "@/lib/anio";

export const dynamic = "force-dynamic";

/** Lee la hoja de planificación (borrador) de CAPEX — vacía si nadie la generó todavía
 *  este año. Siempre lee en vivo, igual que /api/capex. */
export async function GET() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const [proyectos, anioActivo] = await Promise.all([leerBorradorCapex(config, archivo), anioActivoCapex()]);
    return NextResponse.json({
      proyectos,
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
