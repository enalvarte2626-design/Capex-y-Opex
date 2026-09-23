import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { importarBorradorOpex } from "@/lib/planificacionOpex";

export const dynamic = "force-dynamic";

/** Reemplaza todo el borrador de planificación de OPEX con lo que traiga un Excel
 *  subido a mano (mismo layout que "Presupuesto 2026", con Proyectado/Real mes a mes) —
 *  se manda como multipart/form-data, campo "archivo". */
export async function POST(req: NextRequest) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const form = await req.formData().catch(() => null);
  const archivoSubido = form?.get("archivo");
  if (!archivoSubido || !(archivoSubido instanceof Blob)) {
    return NextResponse.json({ error: "Falta el archivo Excel a subir." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const contenido = Buffer.from(await archivoSubido.arrayBuffer());
    const resultado = await importarBorradorOpex(config, archivo, hoja, contenido);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
