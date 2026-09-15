import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, escribirCelda, obtenerConfiguracionOpex, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { resolverCampoOpex, validarValor } from "@/lib/opex-editable";
import { HOJA_PLANIFICACION_OPEX } from "@/lib/planificacionOpex";

export const dynamic = "force-dynamic";

/** Igual que /api/opex/celda, pero escribe en la hoja de planificación (borrador) en
 *  vez de Presupuesto 2026. */
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
    await escribirCelda(config, archivo, HOJA_PLANIFICACION_OPEX, `${definicion.columna}${fila}`, valor);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
