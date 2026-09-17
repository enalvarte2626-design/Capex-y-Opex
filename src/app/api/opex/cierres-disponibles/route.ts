import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { listarArchivosParaComparar } from "@/lib/compararCierre";
import { leerMesCierre } from "@/lib/mesCierreConfig";

export const dynamic = "force-dynamic";

/** Igual que /api/capex/cierres-disponibles, pero para OPEX. La sugerencia de "meses
 *  cerrados" del archivo en vivo no se puede adivinar por el nombre del archivo (OPEX no
 *  sigue el patrón "N+M" hasta que se genera el primer cierre) — se usa el marcador real
 *  guardado dentro del archivo (mismo que ya usa /api/opex/mes-cierre). */
export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const disponibles = await listarArchivosParaComparar(config, archivo);
    const mesCierreActual = await leerMesCierre(config, archivo);
    const disponiblesAjustados = disponibles.map((a) =>
      a.esArchivoActual ? { ...a, cerradosSugeridos: mesCierreActual } : a
    );
    return NextResponse.json({ disponibles: disponiblesAjustados, archivoActual: archivo.nombre });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
