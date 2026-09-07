import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  escribirCelda,
  escribirFila,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import {
  COL_FACTURAS,
  ENCABEZADOS_NUEVOS_FACTURAS,
  extraerFacturas,
  leerCeldaCruda,
  leerWorkbook,
  mesDesdeComentario,
} from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

/** Quita el viejo prefijo "Periodo {Mes}" (con o sin " — algo" después) de un comentario,
 *  dejando solo el resto del texto (o vacío, si no había nada más). */
function comentarioSinPeriodo(comentarios: string): string {
  return comentarios.replace(/^Periodo\s+[A-Za-zÀ-ÿ]+\s*(?:—\s*)?/i, "").trim();
}

/**
 * Migración de datos, un solo uso (idempotente: correrla de nuevo no hace nada si ya no
 * queda ningún "Periodo X" pendiente): para cada factura ya registrada cuyo Mes al que
 * pertenece el gasto todavía solo viva como texto "Periodo {Mes}" dentro de Comentarios,
 * copia ese mes a la nueva columna "Mes Real" y limpia el comentario, dejando solo el
 * texto adicional real (si había).
 *
 * A propósito NO toca el Gasto Real de BD_CAPEX — esos montos ya están cargados en el
 * presupuesto tal como se registraron; esto es solo corregir cómo se ve/identifica el
 * dato, no recalcular nada.
 */
export async function POST() {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    if (!leerCeldaCruda(wb, HOJA_FACTURAS, "J1")) {
      await escribirFila(config, archivo, HOJA_FACTURAS, 1, "J", "N", ENCABEZADOS_NUEVOS_FACTURAS);
    }

    const facturas = extraerFacturas(wb, HOJA_FACTURAS);
    let migradas = 0;
    const sinMes: number[] = [];

    for (const f of facturas) {
      const celdaMes = `${columnaALetra(COL_FACTURAS.mesReal)}${f.filaExcel}`;
      const yaTieneMesReal = Boolean(leerCeldaCruda(wb, HOJA_FACTURAS, celdaMes));
      if (yaTieneMesReal) continue;

      const mes = mesDesdeComentario(f.comentarios);
      if (mes == null) {
        sinMes.push(f.filaExcel);
        continue;
      }

      await escribirCelda(config, archivo, HOJA_FACTURAS, celdaMes, mes);
      const celdaComentario = `${columnaALetra(COL_FACTURAS.comentarios)}${f.filaExcel}`;
      await escribirCelda(config, archivo, HOJA_FACTURAS, celdaComentario, comentarioSinPeriodo(f.comentarios));
      migradas++;
    }

    return NextResponse.json({
      ok: true,
      totalFacturas: facturas.length,
      migradas,
      // Filas que no tenían "Periodo X" en Comentarios ni Mes Real ya puesto — quedan
      // con Mes vacío, corregibles a mano desde la tabla (columna "Mes").
      sinMesDetectado: sinMes,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
