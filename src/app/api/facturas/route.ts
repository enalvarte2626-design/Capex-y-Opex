import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { extraerFacturas, extraerProyectos, leerWorkbook } from "@/lib/capex-parse";
import { resolverFacturaABDCapex } from "@/lib/capex";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

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

    const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const proyectos = extraerProyectos(wb, hojaProyectos);
    const facturas = extraerFacturas(wb, HOJA_FACTURAS)
      .sort((a, b) => b.filaExcel - a.filaExcel)
      .map((f) => ({ ...f, resolucion: resolverFacturaABDCapex(f, proyectos) }));

    return NextResponse.json({
      proyectos: proyectos.map((p) => ({
        filaExcel: p.filaExcel,
        proyecto: p.proyecto,
        detalle: p.detalle,
        grupoNegocio: p.grupoNegocio,
        responsable: p.responsable,
      })),
      facturas,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}