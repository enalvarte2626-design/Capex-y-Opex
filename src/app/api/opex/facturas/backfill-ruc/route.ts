import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  descargarContenido,
  escribirCelda,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_FACTURAS_OPEX, extraerFacturasOpex } from "@/lib/opex-parse";
import { leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

/**
 * RUC conocido por nombre de proveedor — solo aplica a proveedores PERUANOS (el RUC es
 * un identificador tributario de Perú; un proveedor extranjero que factura desde el
 * exterior, como GoDaddy, Adobe o Network Solutions, legítimamente no tiene uno).
 *
 * Investigado una sola vez a mano (búsqueda web) para las facturas ya registradas antes
 * de que existiera el campo RUC en el formulario — de acá en adelante, cada factura
 * nueva ya pide el RUC directo al registrarla, así que este mapa no necesita mantenerse.
 *
 * Comparación por nombre en minúsculas, sin acentos ni espacios de más, para que
 * variaciones de escritura ("Red Cientifica Peruana", "RED CIENTÍFICA PERUANA") igual
 * calcen.
 */
const RUC_CONOCIDO_POR_PROVEEDOR: Record<string, string> = {
  "red cientifica peruana": "20111451592",
};

/** Quita acentos y normaliza espacios/mayúsculas para comparar nombres de proveedor. */
function normalizar(texto: string): string {
  return texto
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "") // marcas diacríticas combinables (acentos)
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Rellena el RUC de facturas ya registradas cuyo proveedor coincide con uno de la lista
 * de arriba y que todavía no tienen RUC guardado. No toca nada más de la fila (monto,
 * mes, etc.) — solo esa celda, y solo si está vacía.
 */
export async function POST() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hojaFacturas = process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const facturas = extraerFacturasOpex(wb, hojaFacturas);

    const actualizadas: Array<{ fila: number; proveedor: string; ruc: string }> = [];

    for (const factura of facturas) {
      if (factura.ruc) continue; // ya tiene RUC, no se toca
      const ruc = RUC_CONOCIDO_POR_PROVEEDOR[normalizar(factura.proveedor)];
      if (!ruc) continue; // proveedor no está en la lista (probablemente extranjero)

      const direccion = `${columnaALetra(COL_FACTURAS_OPEX.ruc)}${factura.filaExcel}`;
      await escribirCelda(config, archivo, hojaFacturas, direccion, ruc);
      actualizadas.push({ fila: factura.filaExcel, proveedor: factura.proveedor, ruc });
    }

    return NextResponse.json({
      ok: true,
      actualizadas,
      mensaje:
        actualizadas.length > 0
          ? `Se completó el RUC de ${actualizadas.length} factura(s).`
          : "No había ninguna factura pendiente con un proveedor peruano conocido.",
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
