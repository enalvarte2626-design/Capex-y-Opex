import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
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
// Corrige muchas filas seguidas (una llamada a Graph por fila) — el límite por defecto
// de la función (10s) no alcanza ni para una base de datos mediana; 60s da margen de
// sobra sin necesitar procesar por lotes.
export const maxDuration = 60;

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
      const celdaMesCruda = leerCeldaCruda(wb, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.mesReal)}${f.filaExcel}`);
      const yaTieneMesReal = Boolean(celdaMesCruda);
      // Reconoce un intento anterior cortado a medias (ej. por un timeout): si el Mes Real
      // ya quedó puesto pero el comentario todavía trae el viejo "Periodo X" sin limpiar,
      // igual hay algo que corregir acá.
      const comentarioTienePeriodo = /^Periodo\s+[A-Za-zÀ-ÿ]+/i.test(f.comentarios);
      if (yaTieneMesReal && !comentarioTienePeriodo) continue;

      const mes = yaTieneMesReal ? Number(celdaMesCruda) : mesDesdeComentario(f.comentarios);
      if (mes == null) {
        sinMes.push(f.filaExcel);
        continue;
      }

      // Un solo PATCH por fila (columnas I-N juntas) en vez de dos por separado — la
      // mitad de las llamadas a Graph, la mitad del tiempo. Los valores de Moneda/Monto
      // Soles/Tipo de Cambio/RUC se re-escriben tal cual ya estaban (crudos, sin
      // reformatear) para no arriesgar ningún cambio en esas columnas.
      await escribirFila(config, archivo, HOJA_FACTURAS, f.filaExcel, "I", "N", [
        comentarioSinPeriodo(f.comentarios),
        leerCeldaCruda(wb, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.moneda)}${f.filaExcel}`),
        leerCeldaCruda(wb, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.montoSoles)}${f.filaExcel}`),
        leerCeldaCruda(wb, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.tipoCambio)}${f.filaExcel}`),
        leerCeldaCruda(wb, HOJA_FACTURAS, `${columnaALetra(COL_FACTURAS.ruc)}${f.filaExcel}`),
        mes,
      ]);
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
