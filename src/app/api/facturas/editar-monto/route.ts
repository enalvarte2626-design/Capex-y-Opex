import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  escribirCelda,
  leerCelda,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_BD, COL_FACTURAS, extraerFacturas, extraerProyectos, leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { resolverFacturaABDCapex } from "@/lib/capex";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

/**
 * Corrige el Monto de una factura ya registrada — y ajusta el Gasto Real de BD_CAPEX
 * por la diferencia (nuevo - anterior), nunca reemplaza el total del mes (puede haber
 * otras facturas sumadas ahí). Solo funciona si la factura se puede emparejar con
 * certeza a una fila/mes de BD_CAPEX (mismo texto que el módulo ya escribe al
 * registrar) — si no, se rechaza para no arriesgar a tocar la celda equivocada.
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json(
      { error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` },
      { status: 500 }
    );
  }

  const cuerpo = await request.json().catch(() => null);
  const filaFactura = Number(cuerpo?.filaFactura);
  const montoNuevo = Number(cuerpo?.montoNuevo);

  if (!Number.isInteger(filaFactura) || filaFactura < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }
  if (!Number.isFinite(montoNuevo) || montoNuevo <= 0) {
    return NextResponse.json({ error: "El monto debe ser mayor a 0." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";

    // Lee todo fresco: la factura (para saber su Proyecto/Comentarios/Monto actual) y
    // BD_CAPEX (para emparejar y saber el Gasto Real vigente).
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const factura = extraerFacturas(wb, HOJA_FACTURAS).find((f) => f.filaExcel === filaFactura);
    if (!factura) {
      return NextResponse.json({ error: "No se encontró esa factura." }, { status: 404 });
    }
    const proyectos = extraerProyectos(wb, hojaProyectos);
    const resolucion = resolverFacturaABDCapex(factura, proyectos);
    if (!resolucion) {
      return NextResponse.json(
        { error: "No se puede identificar con certeza a qué proyecto/mes de BD_CAPEX corresponde esta factura, así que no se puede corregir el monto de forma segura." },
        { status: 409 }
      );
    }

    const montoAnterior = factura.monto;
    const delta = montoNuevo - montoAnterior;

    const colReal = columnaALetra(COL_BD.primerMesReal + (resolucion.mes - 1) * 2);
    const direccionReal = `${colReal}${resolucion.filaProyecto}`;
    const gastoRealAnterior = await leerCelda(config, archivo, hojaProyectos, direccionReal);
    const gastoRealNuevo = gastoRealAnterior + delta;
    await escribirCelda(config, archivo, hojaProyectos, direccionReal, gastoRealNuevo);

    await escribirCelda(
      config,
      archivo,
      HOJA_FACTURAS,
      `${columnaALetra(COL_FACTURAS.monto)}${filaFactura}`,
      montoNuevo
    );

    return NextResponse.json({ ok: true, gastoRealAnterior, gastoRealNuevo });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}