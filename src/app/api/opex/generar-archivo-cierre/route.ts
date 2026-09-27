import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { generarArchivoDeCierreOpex } from "@/lib/mesCierreConfig";

export const dynamic = "force-dynamic";

/** Genera el archivo del siguiente cierre de OPEX (copia con el nombre "N+M" que sigue,
 *  ej. "8+4" → "9+3") — igual que "Generar archivo de cierre" en CAPEX. El archivo
 *  actual no se toca. Ver lib/mesCierreConfig.ts#generarArchivoDeCierreOpex. */
export async function POST() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const resultado = await generarArchivoDeCierreOpex(config, archivo);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
