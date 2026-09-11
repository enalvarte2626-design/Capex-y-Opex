import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  descargarContenido,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { extraerFacturasOpex } from "@/lib/opex-parse";
import { leerWorkbook } from "@/lib/capex-parse";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";

export const dynamic = "force-dynamic";

/**
 * Descarga el reporte de facturas OPEX como un .xlsx real (no solo lo que se ve en
 * pantalla) — con todas las columnas ya resueltas (nombre de mes en vez de número,
 * fecha legible, moneda/empresa/RUC visibles) y ORDENADO por línea presupuestal (Línea
 * de Gasto), para que sea fácil de revisar por línea en vez de por orden de registro.
 */
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

    const hojaFacturas = process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";
    const facturas = extraerFacturasOpex(wb, hojaFacturas);

    // Ordenado por línea presupuestal (y dentro de cada línea, por mes) — a diferencia
    // de la tabla en pantalla, que muestra las más recientes primero.
    const ordenadas = [...facturas].sort((a, b) => {
      const porLinea = a.lineaGasto.localeCompare(b.lineaGasto, "es");
      if (porLinea !== 0) return porLinea;
      return (a.mes ?? 0) - (b.mes ?? 0);
    });

    const filas = ordenadas.map((f) => ({
      Fecha: f.fecha || "—",
      Mes: f.mes ? NOMBRES_MES_CIERRE[f.mes - 1] : "—",
      Empresa: f.empresa || "—",
      "Grupo de Gasto": f.grupoGasto || "—",
      "Subgrupo de Gasto": f.subgrupoGasto || "—",
      "Línea de Gasto": f.lineaGasto || "—",
      Proveedor: f.proveedor || "—",
      RUC: f.ruc || "—",
      "N° Comprobante": f.numeroComprobante || "—",
      Moneda: f.moneda === "PEN" ? "Soles" : f.moneda === "EUR" ? "Euros" : f.moneda === "USD" ? "Dólares" : "—",
      "Monto Soles (sin IGV)": f.montoSoles ?? "",
      "Origen Monto Soles": f.montoSoles == null ? "—" : f.montoSolesEsCalculado ? "Calculado (ref.)" : "Ingresado",
      "Tipo de Cambio (S/)": f.tipoCambio ?? "",
      "Monto Euros": f.montoEuros ?? "",
      "Tipo de Cambio (€)": f.tipoCambioEur ?? "",
      "Monto (USD)": f.monto,
      Comentario: f.comentario || "—",
      Registrado: f.registrado || "—",
    }));

    const hoja = XLSX.utils.json_to_sheet(filas);
    // Ancho de columna aproximado por encabezado — sin esto Excel deja todo apretado a 8
    // caracteres y el reporte queda ilegible al abrirlo.
    hoja["!cols"] = Object.keys(filas[0] ?? {}).map((encabezado) => ({
      wch: Math.max(encabezado.length + 2, 12),
    }));

    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Facturas OPEX");
    const buffer = XLSX.write(libro, { type: "buffer", bookType: "xlsx" }) as Buffer;

    const fechaArchivo = new Date().toISOString().slice(0, 10);
    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="facturas-opex-${fechaArchivo}.xlsx"`,
      },
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
