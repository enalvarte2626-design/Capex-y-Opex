import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { listarArchivosCierreDisponibles } from "@/lib/compararCierre";

export const dynamic = "force-dynamic";

/** Lista los archivos de cierre anteriores guardados en la misma carpeta que el archivo
 *  en vivo — para elegir contra cuál comparar en /bd-capex/comparar-cierre. */
export async function GET() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const disponibles = await listarArchivosCierreDisponibles(config, archivo);
    return NextResponse.json({ disponibles, archivoActual: archivo.nombre });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
