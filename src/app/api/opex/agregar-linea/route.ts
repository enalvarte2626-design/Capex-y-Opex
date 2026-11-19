import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  agregarFilaTabla,
  camposFaltantesOpex,
  contarColumnasTabla,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_PPTO_OPEX } from "@/lib/opex-parse";

export const dynamic = "force-dynamic";

interface CuerpoNuevaLinea {
  empresa: string;
  grupoGasto: string;
  subgrupoGasto: string;
  lineaGasto: string;
  moneda?: string;
  detalle?: string;
  responsable?: string;
  presupuestoAprobado: number;
}

/**
 * Agrega una línea de gasto nueva a "Presupuesto 2026" — directo, sin pasar por
 * Planificación. Los 12 meses (Proyectado/Real) arrancan en 0 — se llenan después desde
 * la misma tabla (celda por celda) o registrando facturas en Facturas OPEX. Status
 * arranca en "Activa".
 *
 * Antes este archivo tenía, por error, la lógica de REGISTRAR UNA FACTURA (de cuando se
 * copió como punto de partida) en vez de la de agregar una línea — por eso el formulario
 * "Nueva línea de gasto OPEX" siempre fallaba con "Línea de gasto inválida.": el cuerpo
 * que mandaba (empresa/grupoGasto/subgrupoGasto/lineaGasto/...) no tenía nada que ver
 * con lo que el código de antes esperaba (filaPresupuesto/mes/monto/proveedor...).
 *
 * "Presupuesto 2026" es una Tabla de Excel real (a diferencia de BD_CAPEX, que es un
 * rango suelto) — se agrega la fila con el endpoint nativo de Tablas de Graph
 * (`agregarFilaTabla`, mismo que ya usa Facturas OPEX), dejando sin especificar (`null`)
 * las 3 columnas calculadas del final (Gasto Real 2026/Forecast 2026/Diferencia): al
 * agregar una fila por este endpoint, Excel les pone sola la misma fórmula que ya tiene
 * el resto de filas de la Tabla, en vez de que la app tenga que adivinar el texto exacto
 * de esa fórmula.
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = (await request.json().catch(() => null)) as CuerpoNuevaLinea | null;
  if (!cuerpo) return NextResponse.json({ error: "Cuerpo inválido." }, { status: 400 });

  const empresa = cuerpo.empresa?.trim() ?? "";
  const grupoGasto = cuerpo.grupoGasto?.trim() ?? "";
  const subgrupoGasto = cuerpo.subgrupoGasto?.trim() ?? "";
  const lineaGasto = cuerpo.lineaGasto?.trim() ?? "";
  const moneda = cuerpo.moneda?.trim() || "USD";
  const detalle = cuerpo.detalle?.trim() ?? "";
  const responsable = cuerpo.responsable?.trim() ?? "";
  const presupuestoAprobado = Number(cuerpo.presupuestoAprobado);

  if (!empresa) return NextResponse.json({ error: "Falta la Empresa." }, { status: 400 });
  if (!grupoGasto) return NextResponse.json({ error: "Falta el Grupo de Gasto." }, { status: 400 });
  if (!subgrupoGasto) return NextResponse.json({ error: "Falta el Subgrupo de Gasto." }, { status: 400 });
  if (!lineaGasto) return NextResponse.json({ error: "Línea de gasto inválida." }, { status: 400 });
  if (!Number.isFinite(presupuestoAprobado) || presupuestoAprobado < 0) {
    return NextResponse.json({ error: "El Presupuesto Aprobado debe ser un número válido." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";

    const fila: Array<string | number | null> = new Array(COL_PPTO_OPEX.presupuestoAprobado + 1).fill("");
    fila[COL_PPTO_OPEX.empresa] = empresa;
    fila[COL_PPTO_OPEX.grupoGasto] = grupoGasto;
    fila[COL_PPTO_OPEX.subgrupoGasto] = subgrupoGasto;
    fila[COL_PPTO_OPEX.lineaGasto] = lineaGasto;
    fila[COL_PPTO_OPEX.status] = "Activa";
    fila[COL_PPTO_OPEX.moneda] = moneda;
    fila[COL_PPTO_OPEX.detalle] = detalle;
    fila[COL_PPTO_OPEX.responsable] = responsable;
    for (let m = 0; m < 12; m++) {
      fila[COL_PPTO_OPEX.primerMesProyectado + m * 2] = 0;
      fila[COL_PPTO_OPEX.primerMesReal + m * 2] = 0;
    }
    fila[COL_PPTO_OPEX.presupuestoAprobado] = presupuestoAprobado;
    // Gasto Real 2026 / Forecast 2026: columnas calculadas de la Tabla, al final — se
    // dejan como `null` a propósito (ver comentario de arriba). "Diferencia" NO es una
    // columna propia de esta Tabla (se confirmó probando contra el Excel real: la Tabla
    // tiene 37 columnas, dos más que Presupuesto Aprobado, no tres) — la app la calcula
    // aparte, nunca la lee ni la escribe acá.
    fila.push(null, null);

    // El endpoint rows/add de Graph exige que el array calce EXACTO con el ancho real de
    // la Tabla (ni una columna de más ni de menos) — se completa con `null` lo que haga
    // falta hasta esa cantidad exacta, en vez de adivinar un ancho fijo.
    const columnasReales = await contarColumnasTabla(config, archivo, hoja);
    if (columnasReales < fila.length) {
      throw new ErrorSharePoint(
        `La Tabla de "${hoja}" tiene menos columnas (${columnasReales}) de las que esta app espera escribir (${fila.length}) — avisa a soporte antes de reintentar, puede haber cambiado el orden de columnas en el Excel.`
      );
    }
    while (fila.length < columnasReales) fila.push(null);

    await agregarFilaTabla(config, archivo, hoja, fila);

    return NextResponse.json({ ok: true });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
