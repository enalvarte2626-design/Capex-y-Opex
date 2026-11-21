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
import { COL_BD, COL_FACTURAS, extraerFacturas, extraerProyectos, fechaAExcelSerial, leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { resolverFacturaABDCapex } from "@/lib/capex";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

/** Campos de texto de una factura que siempre se pueden corregir (no afectan Gasto Real). */
const CAMPOS_TEXTO: Record<string, number> = {
  recurso: COL_FACTURAS.recurso,
  proveedor: COL_FACTURAS.proveedor,
  responsable: COL_FACTURAS.responsable,
  numeroFactura: COL_FACTURAS.numeroFactura,
  comentarios: COL_FACTURAS.comentarios,
  ruc: COL_FACTURAS.ruc,
};

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
  const fila = Number(cuerpo?.fila);
  const campo = String(cuerpo?.campo ?? "");
  const valorCrudo = String(cuerpo?.valor ?? "").trim();

  if (!Number.isInteger(fila) || fila < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }

  if (campo === "periodoFacturado") {
    const fecha = new Date(valorCrudo);
    if (Number.isNaN(fecha.getTime())) {
      return NextResponse.json({ error: "Fecha inválida." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      await escribirCelda(
        config,
        archivo,
        HOJA_FACTURAS,
        `${columnaALetra(COL_FACTURAS.periodoFacturado)}${fila}`,
        fechaAExcelSerial(fecha)
      );
      return NextResponse.json({ ok: true });
    } catch (e) {
      const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
      return NextResponse.json({ error: mensaje }, { status: 502 });
    }
  }

  // Monto en Soles sin IGV — ej. se había puesto el monto CON IGV por error. Solo
  // aplica a facturas ingresadas en Soles (si el valor es calculado a partir del USD,
  // este campo ni siquiera se muestra editable). Recalcula el Monto (USD) con el MISMO
  // tipo de cambio ya guardado en esa factura (el que de verdad se usó al registrarla,
  // no el actual — por si cambió desde entonces) y ajusta el Gasto Real de BD_CAPEX por
  // la diferencia, igual que /api/facturas/editar-monto — así no hay que corregir los
  // dos campos por separado ni arriesgarse a que queden descuadrados entre sí.
  if (campo === "montoSoles") {
    const montoSolesNuevo = Number(valorCrudo);
    if (!Number.isFinite(montoSolesNuevo) || montoSolesNuevo <= 0) {
      return NextResponse.json({ error: "El monto en Soles debe ser un número mayor a 0." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";

      const contenido = await descargarContenido(config, archivo);
      const wb = leerWorkbook(contenido);
      const factura = extraerFacturas(wb, HOJA_FACTURAS).find((f) => f.filaExcel === fila);
      if (!factura) {
        return NextResponse.json({ error: "No se encontró esa factura." }, { status: 404 });
      }
      if (!factura.tipoCambio) {
        return NextResponse.json({ error: "Esta factura no tiene guardado el tipo de cambio usado, así que no se puede recalcular el Monto (USD) con certeza." }, { status: 409 });
      }

      const montoNuevo = Math.round((montoSolesNuevo / factura.tipoCambio) * 100) / 100;
      await escribirCelda(config, archivo, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.montoSoles)}${fila}`, montoSolesNuevo);
      await escribirCelda(config, archivo, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.monto)}${fila}`, montoNuevo);

      const proyectos = extraerProyectos(wb, hojaProyectos);
      const resolucion = resolverFacturaABDCapex(factura, proyectos);
      if (!resolucion) {
        // El Monto (USD) y el Soles ya quedaron corregidos arriba — solo no se pudo
        // ajustar el Gasto Real porque no se identifica con certeza a qué fila/mes
        // corresponde (mismo caso que ya bloquea editar el Monto directamente).
        return NextResponse.json({
          ok: true,
          montoNuevo,
          avisoGastoReal: "No se pudo ajustar el Gasto Real de BD_CAPEX (no se identifica con certeza a qué proyecto/mes corresponde esta factura) — revísalo a mano si hace falta.",
        });
      }

      const colReal = columnaALetra(COL_BD.primerMesReal + (resolucion.mes - 1) * 2);
      const direccionReal = `${colReal}${resolucion.filaProyecto}`;
      const delta = montoNuevo - factura.monto;
      const gastoRealAnterior = await leerCelda(config, archivo, hojaProyectos, direccionReal);
      const gastoRealNuevo = gastoRealAnterior + delta;
      await escribirCelda(config, archivo, hojaProyectos, direccionReal, gastoRealNuevo);

      return NextResponse.json({ ok: true, montoNuevo, gastoRealAnterior, gastoRealNuevo });
    } catch (e) {
      const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
      return NextResponse.json({ error: mensaje }, { status: 502 });
    }
  }

  // Mes al que pertenece el gasto en el presupuesto — solo corrige el dato/etiqueta de
  // esta factura, NUNCA recalcula el Gasto Real de BD_CAPEX (ese ajuste, si hiciera
  // falta, se hace aparte con el campo Monto).
  if (campo === "mesReal") {
    const mes = Number(valorCrudo);
    if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
      return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
    }
    try {
      const archivo = await resolverArchivoPorShareUrl(config);
      await escribirCelda(config, archivo, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.mesReal)}${fila}`, mes);
      return NextResponse.json({ ok: true });
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
    await escribirCelda(config, archivo, HOJA_FACTURAS, `${columnaALetra(indiceColumna)}${fila}`, valorCrudo);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}