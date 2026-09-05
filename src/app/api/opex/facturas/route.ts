import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  descargarContenido,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { extraerFacturasOpex, extraerPresupuestoOpex } from "@/lib/opex-parse";
import { leerWorkbook } from "@/lib/capex-parse";

export const dynamic = "force-dynamic";

export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    const hojaPresupuesto = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const hojaFacturas = process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";

    const lineas = extraerPresupuestoOpex(wb, hojaPresupuesto);
    const facturas = extraerFacturasOpex(wb, hojaFacturas).sort((a, b) => b.filaExcel - a.filaExcel);

    return NextResponse.json({
      lineas: lineas.map((l) => ({
        filaExcel: l.filaExcel,
        grupoGasto: l.grupoNegocio,
        subgrupoGasto: l.prioridad,
        lineaGasto: l.proyecto,
        responsable: l.responsable,
      })),
      facturas,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}