import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { listarVersionesDeArchivo } from "@/lib/compararCierre";

export const dynamic = "force-dynamic";

/** Historial de versiones de SharePoint de un archivo puntual (?itemId=...) — para
 *  elegir una versión anterior a que alguien lo siguiera editando después de "cerrarlo". */
export async function GET(req: NextRequest) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const itemId = req.nextUrl.searchParams.get("itemId")?.trim();
  if (!itemId) {
    return NextResponse.json({ error: "Falta indicar de qué archivo." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const versiones = await listarVersionesDeArchivo(config, archivo.driveId, itemId);
    return NextResponse.json({ versiones });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
