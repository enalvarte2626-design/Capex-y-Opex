import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
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

/** El comentario siempre trae "Periodo {Mes}" (se arma así al registrar la factura, ver
 *  /api/facturas/registrar) — se extrae directo de ahí para el reporte, sin depender de
 *  que `resolverFacturaABDCapex` pueda identificar la fila exacta de BD_CAPEX (el "Mes al
 *  que pertenece el gasto" puede ser distinto a la fecha de emisión del comprobante). */
function mesPresupuestalDeComentario(comentarios: string): string {
  const match = comentarios.match(/Periodo\s+([A-Za-zÀ-ÿ]+)/i);
  return match ? match[1] : "—";
}

/**
 * Descarga el reporte de facturas CAPEX como un .xlsx real, ordenado por proyecto (línea
 * presupuestal) — mismo criterio que el reporte de OPEX (/api/opex/facturas/exportar).
 */
export async function GET() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const proyectos = extraerProyectos(wb, hojaProyectos);
    const facturas = extraerFacturas(wb, HOJA_FACTURAS).map((f) => ({
      ...f,
      resolucion: resolverFacturaABDCapex(f, proyectos),
    }));

    const ordenadas = [...facturas].sort((a, b) => a.proyecto.localeCompare(b.proyecto, "es"));

    const filas = ordenadas.map((f) => ({
      "Fecha Emisión": f.periodoFacturado || "—",
      "Mes Presupuestal": mesPresupuestalDeComentario(f.comentarios),
      Proyecto: f.proyecto || "—",
      Proveedor: f.recurso || "—",
      "Empresa (código)": f.proveedor || "—",
      Responsable: f.responsable || "—",
      "Monto (USD)": f.monto,
      "Monto Soles (sin IGV)": f.montoSoles ?? "—",
      "Origen Monto Soles": f.montoSoles == null ? "—" : f.montoSolesEsCalculado ? "Calculado" : "Ingresado",
      "Tipo de Cambio": f.tipoCambio ?? "—",
      "N° Factura": f.numeroFactura || "—",
      RUC: f.ruc || "—",
      Comentarios: f.comentarios || "—",
      Registrado: f.registrado || "—",
    }));

    const hoja = XLSX.utils.json_to_sheet(filas);
    hoja["!cols"] = Object.keys(filas[0] ?? {}).map((encabezado) => ({
      wch: Math.max(encabezado.length + 2, 12),
    }));

    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Facturas CAPEX");
    const buffer = XLSX.write(libro, { type: "buffer", bookType: "xlsx" }) as Buffer;

    const fechaArchivo = new Date().toISOString().slice(0, 10);
    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="facturas-capex-${fechaArchivo}.xlsx"`,
      },
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
