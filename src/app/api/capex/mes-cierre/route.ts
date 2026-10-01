import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantes, obtenerConfiguracionSharePoint, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { leerMesCierre, escribirMesCierre, mesesCerradosPorNombreArchivo } from "@/lib/mesCierreConfig";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";

export const dynamic = "force-dynamic";

/** Mientras el archivo en vivo nunca haya tenido un "Cerrar mes" explícito (hoja "Config
 *  App" inexistente), el punto de partida es UN MES MENOS que lo que sugiere su nombre
 *  — ej. en "...9+3.xlsm" (recién generado desde "8+4" para poder trabajar Septiembre sin
 *  tocar el archivo ya presentado) el mes que sigue abierto es Septiembre (9), así que el
 *  marcador arranca en 8, no en 9. Antes de que existiera este marcador, generar el
 *  archivo siguiente CERRABA el mes de golpe — ahora son dos decisiones separadas. */
function valorPorDefecto(nombreArchivo: string): number {
  return Math.max(0, mesesCerradosPorNombreArchivo(nombreArchivo) - 1);
}

export async function GET() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const mesCierre = await leerMesCierre(config, archivo, valorPorDefecto(archivo.nombre));
    return NextResponse.json({ mesCierre, nombreMesCierre: NOMBRES_MES_CIERRE[mesCierre - 1] ?? null });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}

/**
 * Cierra un mes más — solo se puede AVANZAR (nunca reabrir uno ya cerrado desde acá), y
 * solo de a un mes a la vez. Separado a propósito de "Generar archivo de cierre"
 * (/api/capex/cerrar-mes, que solo crea el archivo nuevo): generar el archivo ya no
 * cierra nada por sí solo — hay que venir acá explícitamente cuando el mes esté
 * realmente completo.
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
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
    const mesCierreActual = await leerMesCierre(config, archivo, valorPorDefecto(archivo.nombre));

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
    return NextResponse.json({ ok: true, mesCierre: mesNuevo, nombreMesCierre: NOMBRES_MES_CIERRE[mesNuevo - 1] });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
