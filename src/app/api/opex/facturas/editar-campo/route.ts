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
import { fechaAExcelSerial, leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = () => process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";

/** Campos de texto de una factura que siempre se pueden corregir (no afectan Gasto Real). */
const CAMPOS_TEXTO: Record<string, number> = {
  proveedor: COL_FACTURAS_OPEX.proveedor,
  numeroComprobante: COL_FACTURAS_OPEX.numeroComprobante,
  comentario: COL_FACTURAS_OPEX.comentario,
  ruc: COL_FACTURAS_OPEX.ruc,
};

export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = await request.json().catch(() => null);
  const fila = Number(cuerpo?.fila);
  const campo = String(cuerpo?.campo ?? "");
  const valorCrudo = String(cuerpo?.valor ?? "").trim();

  if (!Number.isInteger(fila) || fila < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }

  if (campo === "fecha") {
    const fecha = new Date(valorCrudo);
    if (Number.isNaN(fecha.getTime())) {
      return NextResponse.json({ error: "Fecha inválida." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      await escribirCelda(config, archivo, HOJA_FACTURAS(), `${columnaALetra(COL_FACTURAS_OPEX.fecha)}${fila}`, fechaAExcelSerial(fecha));
      return NextResponse.json({ ok: true });
    } catch (e) {
      const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
      return NextResponse.json({ error: mensaje }, { status: 502 });
    }
  }

  // Mes al que pertenece el gasto en el presupuesto — solo corrige el dato/etiqueta de
  // esta factura, NUNCA recalcula el Gasto Real de Presupuesto 2026 (ese ajuste, si
  // hiciera falta, se hace aparte con el campo Monto).
  if (campo === "mes") {
    const mes = Number(valorCrudo);
    if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
      return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      await escribirCelda(config, archivo, HOJA_FACTURAS(), `${columnaALetra(COL_FACTURAS_OPEX.mes)}${fila}`, mes);
      return NextResponse.json({ ok: true });
    } catch (e) {
      const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
      return NextResponse.json({ error: mensaje }, { status: 502 });
    }
  }

  // Monto en Soles sin IGV — ej. se había puesto el monto CON IGV por error. Solo
  // aplica a facturas ingresadas en Soles (si el valor es calculado a partir del USD,
  // este campo ni siquiera se muestra editable). Recalcula el Monto (USD) con el MISMO
  // tipo de cambio ya guardado en esa factura (el que de verdad se usó al registrarla)
  // y ajusta el Gasto Real de Presupuesto 2026 por la diferencia, igual que
  // /api/opex/facturas/editar-monto — así no hay que corregir los dos campos por
  // separado ni arriesgarse a que queden descuadrados entre sí.
  if (campo === "montoSoles") {
    const montoSolesNuevo = Number(valorCrudo);
    if (!Number.isFinite(montoSolesNuevo) || montoSolesNuevo <= 0) {
      return NextResponse.json({ error: "El monto en Soles debe ser un número mayor a 0." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      const hojaPresupuesto = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
      const hojaFacturas = HOJA_FACTURAS();

      const contenido = await descargarContenido(config, archivo);
      const wb = leerWorkbook(contenido);
      const factura = extraerFacturasOpex(wb, hojaFacturas).find((f) => f.filaExcel === fila);
      if (!factura) {
        return NextResponse.json({ error: "No se encontró esa factura." }, { status: 404 });
      }
      if (!factura.tipoCambio) {
        return NextResponse.json({ error: "Esta factura no tiene guardado el tipo de cambio usado, así que no se puede recalcular el Monto (USD) con certeza." }, { status: 409 });
      }

      const montoNuevo = Math.round((montoSolesNuevo / factura.tipoCambio) * 100) / 100;
      await escribirCelda(config, archivo, hojaFacturas, `${columnaALetra(COL_FACTURAS_OPEX.montoSoles)}${fila}`, montoSolesNuevo);
      await escribirCelda(config, archivo, hojaFacturas, `${columnaALetra(COL_FACTURAS_OPEX.monto)}${fila}`, montoNuevo);

      if (!factura.filaPresupuesto || !factura.mes) {
        // El Monto (USD) y el Soles ya quedaron corregidos arriba — solo no se pudo
        // ajustar el Gasto Real porque esta factura no tiene guardada la línea/mes.
        return NextResponse.json({
          ok: true,
          montoNuevo,
          avisoGastoReal: "No se pudo ajustar el Gasto Real de Presupuesto 2026 (esta factura no tiene guardada la línea de gasto o el mes) — revísalo a mano si hace falta.",
        });
      }

      const colReal = columnaALetra(COL_PPTO_OPEX.primerMesReal + (factura.mes - 1) * 2);
      const direccionReal = `${colReal}${factura.filaPresupuesto}`;
      const delta = montoNuevo - factura.monto;
      const gastoRealAnterior = await leerCelda(config, archivo, hojaPresupuesto, direccionReal);
      const gastoRealNuevo = gastoRealAnterior + delta;
      await escribirCelda(config, archivo, hojaPresupuesto, direccionReal, gastoRealNuevo);

      return NextResponse.json({ ok: true, montoNuevo, gastoRealAnterior, gastoRealNuevo });
    } catch (e) {
      const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
      return NextResponse.json({ error: mensaje }, { status: 502 });
    }
  }

  const indiceColumna = CAMPOS_TEXTO[campo];
  if (indiceColumna === undefined) {
    return NextResponse.json({ error: `Campo "${campo}" no es editable.` }, { status: 400 });
  }
  if (valorCrudo.length > 300) {
    return NextResponse.json({ error: "Texto demasiado largo." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    await escribirCelda(config, archivo, HOJA_FACTURAS(), `${columnaALetra(indiceColumna)}${fila}`, valorCrudo);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}