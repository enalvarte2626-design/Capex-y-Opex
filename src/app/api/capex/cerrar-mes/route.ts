import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  crearArchivo,
  descargarContenido,
  escribirColumna,
  listarNombresCarpeta,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_BD, extraerProyectos, leerCeldaCruda, leerWorkbook, ultimaFilaConDatos } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

/** Convierte "Control Capex Forecast 7+5.xlsm" → { cerrados: 7, restantes: 5 }. */
function leerPatronNombre(nombre: string): { cerrados: number; restantes: number; match: string } | null {
  const m = nombre.match(/(\d+)\s*\+\s*(\d+)/);
  if (!m) return null;
  return { cerrados: Number(m[1]), restantes: Number(m[2]), match: m[0] };
}

/**
 * Genera el archivo del siguiente mes de cierre: una copia exacta del actual, con la
 * única diferencia real de que la fórmula FORESCAST (columna AN) en BD_CAPEX ahora suma
 * un mes proyectado menos (el que se acaba de cerrar). El archivo actual no se toca —
 * esto solo crea uno nuevo al lado, con el nombre que sigue en la numeración
 * ("7+5" → "8+4"), igual que ya se hace a mano cada mes.
 */
export async function POST() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json(
      { error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` },
      { status: 500 }
    );
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const patron = leerPatronNombre(archivo.nombre);
    if (!patron) {
      return NextResponse.json(
        { error: `El nombre del archivo actual ("${archivo.nombre}") no sigue el patrón "N+M" esperado.` },
        { status: 400 }
      );
    }
    if (patron.restantes < 1) {
      return NextResponse.json({ error: "Ya no quedan meses por cerrar este año." }, { status: 400 });
    }

    const nuevoCerrados = patron.cerrados + 1;
    const nuevoRestantes = patron.restantes - 1;
    const nuevoNombre = archivo.nombre.replace(patron.match, `${nuevoCerrados}+${nuevoRestantes}`);

    const existentes = await listarNombresCarpeta(config, archivo.driveId, archivo.carpetaId);
    if (existentes.some((n) => n.toLowerCase() === nuevoNombre.toLowerCase())) {
      return NextResponse.json({ error: `Ya existe un archivo llamado "${nuevoNombre}" en esa carpeta.` }, { status: 409 });
    }

    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const proyectos = extraerProyectos(wb, hoja);
    const filasProyecto = new Set(proyectos.map((p) => p.filaExcel));
    const ultimaFila = ultimaFilaConDatos(wb, hoja);

    const colForecast = columnaALetra(39); // AN — dos columnas después de AL (presupuestoAprobado=37)
    // Columnas proyectado de los meses que quedan por delante del nuevo cierre.
    const columnasProyectadoRestantes = Array.from({ length: 12 - nuevoCerrados }, (_, i) => {
      const mes = nuevoCerrados + i;
      return columnaALetra(COL_BD.primerMesReal + mes * 2 + 1);
    });

    const valoresColumna: Array<string | number> = [];
    for (let fila = 2; fila <= ultimaFila; fila++) {
      if (filasProyecto.has(fila)) {
        const formula = `=SUM(${columnasProyectadoRestantes.map((c) => `${c}${fila}`).join("+")})`;
        valoresColumna.push(formula);
      } else {
        // Filas que no son proyectos (vacías, o subtotales sueltos): se preservan tal cual.
        valoresColumna.push(leerCeldaCruda(wb, hoja, `${colForecast}${fila}`));
      }
    }

    const nuevoArchivo = await crearArchivo(config, archivo.driveId, archivo.carpetaId, nuevoNombre, contenido);
    await escribirColumna(config, nuevoArchivo, hoja, colForecast, 2, ultimaFila, valoresColumna);

    return NextResponse.json({ ok: true, archivo: nuevoNombre });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}