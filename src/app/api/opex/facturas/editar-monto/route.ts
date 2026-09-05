import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  descargarContenido,
  escribirCelda,
  leerCelda,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_FACTURAS_OPEX, COL_PPTO_OPEX, extraerFacturasOpex } from "@/lib/opex-parse";
import { leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = () => process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";

/**
 * Corrige el Monto de una factura de OPEX ya registrada — ajusta el Gasto Real de
 * "Presupuesto 2026" por la diferencia (nuevo - anterior), nunca reemplaza el total del
 * mes. A diferencia de CAPEX, aquí no hace falta "adivinar" a qué línea/mes corresponde:
 * cada factura registrada desde la app ya guarda la fila de Presupuesto y el mes de
 * forma directa — solo se rechaza si esos datos faltan (factura muy vieja o corrupta).
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
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
    const hojaPresupuesto = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const hojaFacturas = HOJA_FACTURAS();

    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const factura = extraerFacturasOpex(wb, hojaFacturas).find((f) => f.filaExcel === filaFactura);
    if (!factura) {
      return NextResponse.json({ error: "No se encontró esa factura." }, { status: 404 });
    }
    if (!factura.filaPresupuesto || !factura.mes) {
      return NextResponse.json(
        { error: "Esta factura no tiene guardada la línea de gasto o el mes, así que no se puede corregir el monto de forma segura." },
        { status: 409 }
      );
    }

    const montoAnterior = factura.monto;
    const delta = montoNuevo - montoAnterior;

    const colReal = columnaALetra(COL_PPTO_OPEX.primerMesReal + (factura.mes - 1) * 2);
    const direccionReal = `${colReal}${factura.filaPresupuesto}`;
    const gastoRealAnterior = await leerCelda(config, archivo, hojaPresupuesto, direccionReal);
    const gastoRealNuevo = gastoRealAnterior + delta;
    await escribirCelda(config, archivo, hojaPresupuesto, direccionReal, gastoRealNuevo);

    await escribirCelda(
      config,
      archivo,
      hojaFacturas,
      `${columnaALetra(COL_FACTURAS_OPEX.monto)}${filaFactura}`,
      montoNuevo
    );

    return NextResponse.json({ ok: true, gastoRealAnterior, gastoRealNuevo });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}