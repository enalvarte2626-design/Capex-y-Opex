import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  escribirCelda,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { resolverCampo, validarValor } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

/**
 * Escribe un campo puntual de un proyecto (fila) de BD_CAPEX. Solo acepta los campos
 * listados en `capex-editable.ts` — cualquier otro identificador se rechaza. Nunca
 * toca más de una celda por llamada.
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
  const fila = Number(cuerpo?.fila);
  const campo = String(cuerpo?.campo ?? "");

  if (!Number.isInteger(fila) || fila < 2) {
    return NextResponse.json({ error: "Fila inválida." }, { status: 400 });
  }

  const definicion = resolverCampo(campo);
  if (!definicion) {
    return NextResponse.json({ error: `Campo "${campo}" no es editable.` }, { status: 400 });
  }

  const valor = validarValor(definicion, cuerpo?.valor);
  if (valor === null) {
    return NextResponse.json({ error: "Valor inválido para este campo." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    await escribirCelda(config, archivo, hoja, `${definicion.columna}${fila}`, valor);
    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}