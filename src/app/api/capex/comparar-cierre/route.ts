import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { compararConCierreAnterior } from "@/lib/compararCierre";

export const dynamic = "force-dynamic";

/** Compara BD_CAPEX en vivo contra un archivo de cierre anterior (?archivo=nombre.xlsm)
 *  — ver lib/compararCierre.ts para el detalle de qué cuenta como "cambio". */
export async function GET(req: NextRequest) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const nombreArchivoAnterior = req.nextUrl.searchParams.get("archivo")?.trim();
  if (!nombreArchivoAnterior) {
    return NextResponse.json({ error: "Falta indicar contra qué archivo de cierre comparar." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const resultado = await compararConCierreAnterior(config, archivo, hoja, nombreArchivoAnterior);
    return NextResponse.json(resultado);
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
