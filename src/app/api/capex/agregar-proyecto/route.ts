import { NextRequest, NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  escribirColumna,
  escribirFila,
  insertarFila,
  leerCelda,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { leerCeldaCruda, leerWorkbook, ultimaFilaConDatosEscaneada } from "@/lib/capex-parse";

export const dynamic = "force-dynamic";

/**
 * Agrega un proyecto nuevo a BD_CAPEX. La hoja no tiene una fila vacía "suelta" al
 * final: la última fila con datos es "TOTAL USD", con fórmulas SUBTOTAL que suman desde
 * la fila 2 hasta la última fila de proyecto — y justo antes de esa fila de TOTAL hay
 * una fila en blanco (el hueco que deja el propio Excel). Un proyecto nuevo entra ahí,
 * y las ~28 fórmulas de la fila TOTAL se extienden una fila para incluirlo — si no se
 * hiciera esto, el proyecto se vería en la tabla pero no sumaría en los totales.
 */
export async function POST(req: NextRequest) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const body = await req.json().catch(() => null);
  const proyecto = String(body?.proyecto ?? "").trim();
  const subNegocio = String(body?.subNegocio ?? "").trim();
  const grupoNegocio = String(body?.grupoNegocio ?? "").trim().toUpperCase();
  const detalle = String(body?.detalle ?? "").trim();
  const prioridad = String(body?.prioridad ?? "").trim();
  const presupuestoAprobado = Number(body?.presupuestoAprobado);
  // Gasto Real inicial (opcional): un gasto ya ejecutado desde el momento de crear el
  // proyecto, cargado directo al mes indicado (misma columna "Real" que ya se edita en
  // la tabla) — así el Gasto Real no queda en $0 si el proyecto ya arrancó con gasto.
  const gastoRealInicial = body?.gastoRealInicial != null && body?.gastoRealInicial !== "" ? Number(body.gastoRealInicial) : 0;
  const mesGastoReal = body?.mesGastoReal != null && body?.mesGastoReal !== "" ? Number(body.mesGastoReal) : null;
  // Monto Proyectado (opcional): mismo criterio que el Gasto Real inicial, pero cargado
  // en la columna "Proyectado" del mes indicado — para no dejar el Forecast en $0 si ya
  // se sabe cuándo se espera ejecutar el gasto.
  const montoProyectado = body?.montoProyectado != null && body?.montoProyectado !== "" ? Number(body.montoProyectado) : 0;
  const mesProyectado = body?.mesProyectado != null && body?.mesProyectado !== "" ? Number(body.mesProyectado) : null;

  if (!["EMISIVO", "RECEPTIVO", "TRANSVERSAL"].includes(grupoNegocio)) {
    return NextResponse.json({ error: "Grupo de Negocio debe ser Emisivo, Receptivo o Transversal." }, { status: 400 });
  }
  if (!prioridad) return NextResponse.json({ error: "Falta la Prioridad." }, { status: 400 });
  if (!Number.isFinite(presupuestoAprobado) || presupuestoAprobado < 0) {
    return NextResponse.json({ error: "El Presupuesto Aprobado debe ser un número válido." }, { status: 400 });
  }
  if (gastoRealInicial > 0 && (!mesGastoReal || mesGastoReal < 1 || mesGastoReal > 12)) {
    return NextResponse.json({ error: "Si ingresas un Gasto Real inicial, elige a qué mes corresponde." }, { status: 400 });
  }
  if (!Number.isFinite(gastoRealInicial) || gastoRealInicial < 0) {
    return NextResponse.json({ error: "El Gasto Real inicial debe ser un número válido." }, { status: 400 });
  }
  if (montoProyectado > 0 && (!mesProyectado || mesProyectado < 1 || mesProyectado > 12)) {
    return NextResponse.json({ error: "Si ingresas un Monto Proyectado, elige a qué mes corresponde." }, { status: 400 });
  }
  if (!Number.isFinite(montoProyectado) || montoProyectado < 0) {
    return NextResponse.json({ error: "El Monto Proyectado debe ser un número válido." }, { status: 400 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hoja = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    // "filaTotalOriginal" es la posición de TOTAL en el snapshot local recién descargado
    // — se sigue usando para leer el TEXTO de sus fórmulas (no cambia con un insert). Si
    // hace falta insertar una fila, TOTAL se corre hacia abajo en el archivo real, y de
    // ahí en adelante "filaTotal" (la posición vigente) es la que se usa para leer/escribir
    // contra SharePoint.
    const filaTotalOriginal = ultimaFilaConDatosEscaneada(wb, hoja);
    const filaNueva = filaTotalOriginal - 1;
    const filaAnterior = filaNueva - 1; // última fila de proyecto ya existente (ej. 83)
    let filaTotal = filaTotalOriginal;

    // Si esa fila no está realmente vacía (ya no queda ningún hueco libre), se inserta una
    // fila nueva justo ahí — igual que "Insertar fila" en Excel: TOTAL y todo lo de abajo
    // se corre un lugar hacia abajo solo, sin pisar nada.
    const celdaA = leerCeldaCruda(wb, hoja, `A${filaNueva}`);
    if (String(celdaA).trim() !== "") {
      await insertarFila(config, archivo, hoja, filaNueva);
      filaTotal = filaTotalOriginal + 1;
    }

    // 1) Datos del proyecto — A:E (Proyecto, Sub Negocio, Grupo, Detalle, % Avance=0).
    await escribirFila(config, archivo, hoja, filaNueva, "A", "E", [proyecto, subNegocio, grupoNegocio, detalle, 0]);

    // 2) Columna F (texto de Avance): misma fórmula que usa cada fila existente.
    await escribirColumna(config, archivo, hoja, "F", filaNueva, filaNueva, [
      `=IF(E${filaNueva}=0%,"No iniciado",IF(E${filaNueva}<30%,"Inicio",IF(E${filaNueva}<80%,"En proceso",IF(E${filaNueva}<100%,"Por culminar","Culminado"))))`,
    ]);

    // 3) G:AL — Categoría/Status/OPEX/Recurso/Responsable/Tiempo vacíos, 12 meses en 0
    // (salvo el mes de Gasto Real inicial, si se indicó uno), Presupuesto Aprobado.
    const meses = Array(24).fill(0);
    if (gastoRealInicial > 0 && mesGastoReal) {
      meses[(mesGastoReal - 1) * 2] = gastoRealInicial; // columna "Real" de ese mes
    }
    if (montoProyectado > 0 && mesProyectado) {
      meses[(mesProyectado - 1) * 2 + 1] = montoProyectado; // columna "Proyectado" de ese mes
    }
    await escribirFila(config, archivo, hoja, filaNueva, "G", "AL", [
      "",
      prioridad,
      "",
      0,
      "",
      "",
      "",
      ...meses,
      presupuestoAprobado,
    ]);

    // 4) AM/AN/AO (Gasto Real/Forecast/Diferencia): mismas fórmulas que usa cada fila.
    await escribirColumna(config, archivo, hoja, "AM", filaNueva, filaNueva, [
      `=SUM(N${filaNueva}+P${filaNueva}+R${filaNueva}+T${filaNueva}+V${filaNueva}+X${filaNueva}+Z${filaNueva}+AB${filaNueva}+AD${filaNueva}+AF${filaNueva}+AH${filaNueva}+AJ${filaNueva})`,
    ]);
    await escribirColumna(config, archivo, hoja, "AN", filaNueva, filaNueva, [
      `=SUM(AC${filaNueva}+AE${filaNueva}+AG${filaNueva}+AI${filaNueva}+AK${filaNueva})`,
    ]);
    await escribirColumna(config, archivo, hoja, "AO", filaNueva, filaNueva, [
      `=AL${filaNueva}-AM${filaNueva}-AN${filaNueva}`,
    ]);

    // 5) Extiende las fórmulas SUBTOTAL de la fila TOTAL (N..AO) para que incluyan la
    // fila nueva — si no, el proyecto no sumaría en los totales.
    const columnasTotal = [
      "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
      "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
      "AL", "AM", "AN", "AO",
    ];
    for (const col of columnasTotal) {
      // El texto de la fórmula se lee del snapshot local (posición original, sin
      // desplazar) — pero se escribe en "filaTotal", la posición vigente en el archivo
      // real (ya corrida si hizo falta insertar una fila).
      const formulaActual = String(leerCeldaCruda(wb, hoja, `${col}${filaTotalOriginal}`));
      if (!formulaActual.startsWith("=")) continue; // por si alguna quedó vacía, no tocarla
      const formulaExtendida = formulaActual.replace(new RegExp(`\\b${filaAnterior}\\b`, "g"), String(filaNueva));
      await escribirColumna(config, archivo, hoja, col, filaTotal, filaTotal, [formulaExtendida]);
    }

    // Confirma releyendo el Presupuesto Aprobado recién escrito, como chequeo simple de
    // que la escritura llegó de verdad (no solo que Graph respondió 200).
    const verificacion = await leerCelda(config, archivo, hoja, `AL${filaNueva}`);

    return NextResponse.json({
      ok: true,
      fila: filaNueva,
      presupuestoAprobadoVerificado: verificacion,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}