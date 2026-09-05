import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { extraerProyeccionVM, extraerProyectos, leerWorkbook } from "@/lib/capex-parse";

export const dynamic = "force-dynamic";

/**
 * Trae siempre los datos más recientes del Excel: el dashboard se refresca al abrir
 * (o al presionar "Actualizar"), no hay caché ni base de datos propia.
 */
export async function GET() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json(
      { error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` },
      { status: 500 }
    );
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    const hojaActual = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const hojaProyeccion = process.env.SP_CAPEX_HOJA_PROYECCION?.trim() || "Proyeccion 2026_vm";

    const proyectos = extraerProyectos(wb, hojaActual);
    const proyeccion = extraerProyeccionVM(wb, hojaProyeccion);

    return NextResponse.json({
      proyectos,
      proyeccion,
      archivo: archivo.nombre,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}