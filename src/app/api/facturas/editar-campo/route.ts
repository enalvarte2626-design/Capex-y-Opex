import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  escribirCelda,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_FACTURAS, fechaAExcelSerial } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

/** Campos de texto de una factura que siempre se pueden corregir (no afectan Gasto Real). */
const CAMPOS_TEXTO: Record<string, number> = {
  recurso: COL_FACTURAS.recurso,
  proveedor: COL_FACTURAS.proveedor,
  responsable: COL_FACTURAS.responsable,
  numeroFactura: COL_FACTURAS.numeroFactura,
  comentarios: COL_FACTURAS.comentarios,
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