import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  escribirCelda,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { resolverCampoOpex, validarValor } from "@/lib/opex-editable";

export const dynamic = "force-dynamic";

/** Igual que /api/capex/celda pero para "Presupuesto 2026" — solo acepta los campos
 *  listados en opex-editable.ts. */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = await request.json().catch(() => null);
  const fila = Number(cuerpo?.fila);
  const campo = String(cuerpo?.campo ?? "");

  if (!Number.isInteger(fila) || fila < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }

  const definicion = resolverCampoOpex(campo);
  if (!definicion) {
    return NextResponse.json({ error: `Campo "${campo}" no es editable.` }, { status: 400 });
  }

  const valor = validarValor(definicion, cuerpo?.valor);
  if (valor === null) {
    return NextResponse.json({ error: "Valor inválido para este campo." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    await escribirCelda(config, archivo, hoja, `${definicion.columna}${fila}`, valor);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}