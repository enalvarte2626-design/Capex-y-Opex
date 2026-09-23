import { NextRequest, NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { importarBorradorCapex } from "@/lib/planificacionCapex";

export const dynamic = "force-dynamic";

/** Reemplaza todo el borrador de planificación de CAPEX con lo que traiga un Excel
 *  subido a mano (mismo layout que BD_CAPEX, con Real/Proyectado mes a mes) — se manda
 *  como multipart/form-data, campo "archivo". */
export async function POST(req: NextRequest) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
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
    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const contenido = Buffer.from(await archivoSubido.arrayBuffer());
    const resultado = await importarBorradorCapex(config, archivo, hoja, contenido);
    return NextResponse.json({ ok: true, ...resultado });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
