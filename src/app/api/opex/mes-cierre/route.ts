import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { leerMesCierre, escribirMesCierre } from "@/lib/mesCierreConfig";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";

export const dynamic = "force-dynamic";

export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const mesCierre = await leerMesCierre(config, archivo);
    return NextResponse.json({ mesCierre, nombreMesCierre: NOMBRES_MES_CIERRE[mesCierre - 1] });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}

/**
 * Cierra un mes más — solo se puede AVANZAR (nunca reabrir uno ya cerrado desde acá), y
 * solo de a un mes a la vez, para que no se salte un mes por error: cerrar Agosto exige
 * que el mes de cierre actual sea Julio.
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = await request.json().catch(() => null);
  const mesNuevo = Number(cuerpo?.mes);
  if (!Number.isInteger(mesNuevo) || mesNuevo < 1 || mesNuevo > 12) {
    return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const mesCierreActual = await leerMesCierre(config, archivo);

    if (mesNuevo <= mesCierreActual) {
      return NextResponse.json(
        { error: `${NOMBRES_MES_CIERRE[mesNuevo - 1]} ya está cerrado (o antes del cierre actual).` },
        { status: 400 }
      );
    }
    if (mesNuevo > mesCierreActual + 1) {
      return NextResponse.json(
        {
          error: `No se puede saltar directo a ${NOMBRES_MES_CIERRE[mesNuevo - 1]} — primero hay que cerrar ${NOMBRES_MES_CIERRE[mesCierreActual]}.`,
        },
        { status: 400 }
      );
    }

    await escribirMesCierre(config, archivo, mesNuevo);
    return NextResponse.json({
      ok: true,
      mesCierre: mesNuevo,
      nombreMesCierre: NOMBRES_MES_CIERRE[mesNuevo - 1],
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
