import { NextResponse } from "next/server";
import { ErrorSharePoint, camposFaltantesOpex, obtenerConfiguracionOpex, renombrarArchivo, resolverArchivoPorShareUrl } from "@/lib/sharepoint";
import { construirNombreCierre, leerMesCierre } from "@/lib/mesCierreConfig";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";

export const dynamic = "force-dynamic";

/**
 * Renombra el archivo en vivo de OPEX al patrón del PRÓXIMO cierre (mesCierreActual+1),
 * SIN avanzar el mes de cierre todavía — para poder empezar a trabajar el mes siguiente
 * (registrar sus facturas, que ya suman solas al Gasto Real porque no está "cerrado")
 * sin seguir escribiendo sobre un archivo cuyo nombre implica que ya se presentó (ej.
 * "...8+4.xlsx"). Mismo criterio que "Generar archivo de cierre" en CAPEX, adaptado a
 * que OPEX no crea copias: solo cambia el nombre del mismo archivo.
 *
 * Cuando más adelante se cierre ese mes de verdad desde "Cerrar mes", el nombre ya va a
 * coincidir con el que corresponde, así que ese paso solo actualiza el marcador — no
 * hace falta renombrar dos veces.
 */
export async function POST() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const mesCierreActual = await leerMesCierre(config, archivo);
    if (mesCierreActual >= 12) {
      return NextResponse.json({ error: "Ya no quedan meses por cerrar este año." }, { status: 400 });
    }

    const proximoMes = mesCierreActual + 1;
    const nuevoNombre = construirNombreCierre(archivo.nombre, proximoMes);
    if (nuevoNombre.toLowerCase() === archivo.nombre.toLowerCase()) {
      return NextResponse.json({ error: `El archivo ya se llama "${nuevoNombre}".` }, { status: 400 });
    }

    await renombrarArchivo(config, archivo.driveId, archivo.itemId, nuevoNombre);
    return NextResponse.json({
      ok: true,
      archivo: nuevoNombre,
      proximoMes,
      nombreProximoMes: NOMBRES_MES_CIERRE[proximoMes - 1],
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
