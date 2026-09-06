import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  agregarFilaTabla,
  asegurarTablaEnHoja,
  camposFaltantesOpex,
  crearHojaSiNoExiste,
  descargarContenido,
  escribirCelda,
  escribirFila,
  leerCelda,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import {
  COL_PPTO_OPEX,
  ENCABEZADOS_FACTURAS_OPEX,
  extraerFacturasOpex,
  extraerPresupuestoOpex,
} from "@/lib/opex-parse";
import { fechaAExcelSerial, leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/opex-constantes";

export const dynamic = "force-dynamic";

export async function GET() {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);

    const hojaPresupuesto = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const hojaFacturas = process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";

    const lineas = extraerPresupuestoOpex(wb, hojaPresupuesto);
    const facturas = extraerFacturasOpex(wb, hojaFacturas).sort((a, b) => b.filaExcel - a.filaExcel);

    return NextResponse.json({
      lineas: lineas.map((l) => ({
        filaExcel: l.filaExcel,
        grupoGasto: l.grupoNegocio,
        subgrupoGasto: l.prioridad,
        lineaGasto: l.proyecto,
        responsable: l.responsable,
      })),
      facturas,
      actualizadoEn: new Date().toISOString(),
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}

interface CuerpoRegistro {
  filaPresupuesto: number;
  mes: number; // 1-12
  /** Monto en Soles SIN IGV que ingresa la persona — el USD que de verdad mueve el
   *  presupuesto se calcula acá abajo (montoSoles / TIPO_CAMBIO_POR_DEFECTO), nunca se
   *  recibe directo del formulario. */
  montoSoles: number;
  proveedor: string;
  numeroComprobante: string;
  comentario?: string;
}

/**
 * Registra una factura de OPEX desde cero (la hoja "Facturas Opex - App" no existía
 * antes de este módulo — se crea sola la primera vez que hace falta, con encabezados).
 * La línea de gasto se elige de antemano (Grupo→Subgrupo→Línea, ya resuelta a una fila
 * puntual de "Presupuesto 2026") — nunca se adivina por texto. Suma el monto al Gasto
 * Real del mes correspondiente sin reemplazar lo que ya hubiera (una línea puede tener
 * varias facturas en el mismo mes).
 *
 * El monto se ingresa en Soles SIN IGV (así llegan la mayoría de las facturas locales) y
 * se convierte acá a USD con el tipo de cambio fijo de la app (`TIPO_CAMBIO_POR_DEFECTO`,
 * hoy 3.4) — ese USD es el que de verdad se suma al Gasto Real; nunca se confía en un
 * monto en USD calculado del lado del navegador.
 *
 * NOTA: esta misma lógica también vive en /api/opex/agregar-linea — se mantuvieron
 * ambas rutas equivalentes a propósito porque no fue posible confirmar cuál de las dos
 * llama de verdad el formulario en producción (verificar en la pestaña Red del
 * navegador y, una vez confirmado, eliminar la ruta que no se use).
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = (await request.json().catch(() => null)) as CuerpoRegistro | null;
  if (!cuerpo) return NextResponse.json({ error: "Cuerpo inválido." }, { status: 400 });

  const { filaPresupuesto, mes, montoSoles, proveedor, numeroComprobante, comentario } = cuerpo;

  if (!Number.isInteger(filaPresupuesto) || filaPresupuesto < 2) {
    return NextResponse.json({ error: "Línea de gasto inválida." }, { status: 400 });
  }
  if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
    return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
  }
  // Un mes ya pasado sí se puede registrar (queda en el historial de "Facturas Opex -
  // App"), pero NUNCA suma al Gasto Real de Presupuesto 2026 — ese presupuesto ya se dio
  // por cerrado para ese mes. Solo el mes actual o uno futuro mueven el presupuesto.
  const mesActualReal = new Date().getMonth() + 1;
  const esMesPasado = mes < mesActualReal;
  if (!Number.isFinite(montoSoles) || montoSoles <= 0) {
    return NextResponse.json({ error: "El monto en Soles debe ser mayor a 0." }, { status: 400 });
  }
  if (!proveedor?.trim()) {
    return NextResponse.json({ error: "Falta el Proveedor." }, { status: 400 });
  }

  const tipoCambio = TIPO_CAMBIO_POR_DEFECTO;
  if (!Number.isFinite(tipoCambio) || tipoCambio <= 0) {
    // No debería pasar nunca (es una constante fija), pero si algún día vuelve a
    // resolverse mal en el bundle del servidor, mejor fallar fuerte acá que guardar un
    // Monto (USD) en null/0 sin que nadie se entere.
    return NextResponse.json({ error: "Tipo de cambio inválido en el servidor." }, { status: 500 });
  }
  const monto = Math.round((montoSoles / tipoCambio) * 100) / 100;

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hojaPresupuesto = process.env.SP_OPEX_HOJA_PRESUPUESTO?.trim() || "Presupuesto 2026";
    const hojaFacturas = process.env.SP_OPEX_HOJA_FACTURAS?.trim() || "Facturas Opex - App";

    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const lineas = extraerPresupuestoOpex(wb, hojaPresupuesto);
    const linea = lineas.find((l) => l.filaExcel === filaPresupuesto);
    if (!linea) {
      return NextResponse.json({ error: "No se encontró esa línea de gasto en Presupuesto 2026." }, { status: 404 });
    }

    await crearHojaSiNoExiste(config, archivo, hojaFacturas);
    const contenidoActualizado = await descargarContenido(config, archivo);
    const wbActualizado = leerWorkbook(contenidoActualizado);
    const hojaExisteConDatos = wbActualizado.Sheets[hojaFacturas]?.["!ref"] != null;
    if (!hojaExisteConDatos) {
      await escribirFila(config, archivo, hojaFacturas, 1, "A", "O", ENCABEZADOS_FACTURAS_OPEX);
    }
    // Asegura que la hoja tenga una Tabla de Excel real antes de agregar la fila — ver
    // el comentario de `agregarFilaTabla` en sharepoint.ts sobre por qué esto reemplaza
    // el cálculo manual de "última fila + 1" (tenía una condición de carrera real: dos
    // facturas registradas seguidas podían pisarse entre sí en la misma fila).
    await asegurarTablaEnHoja(config, archivo, hojaFacturas, "A1:O1");
    await agregarFilaTabla(config, archivo, hojaFacturas, [
      fechaAExcelSerial(new Date()),
      linea.grupoNegocio,
      linea.prioridad,
      linea.proyecto,
      filaPresupuesto,
      mes,
      monto,
      proveedor.trim(),
      numeroComprobante?.trim() ?? "",
      linea.responsable,
      comentario?.trim() ?? "",
      esMesPasado ? "ok (mes pasado, no afecta presupuesto)" : "ok",
      montoSoles,
      tipoCambio,
      linea.subNegocio,
    ]);

    // Un mes pasado queda solo en el historial — nunca toca Presupuesto 2026.
    if (esMesPasado) {
      return NextResponse.json({
        ok: true,
        monto,
        montoSoles,
        tipoCambio,
        presupuestoActualizado: false,
        aviso: "Mes pasado: la factura quedó registrada en el historial, pero no se sumó al Gasto Real de Presupuesto 2026.",
      });
    }

    const colReal = columnaALetra(COL_PPTO_OPEX.primerMesReal + (mes - 1) * 2);
    const direccionReal = `${colReal}${filaPresupuesto}`;
    const valorActual = await leerCelda(config, archivo, hojaPresupuesto, direccionReal);
    const nuevoValor = valorActual + monto;
    await escribirCelda(config, archivo, hojaPresupuesto, direccionReal, nuevoValor);

    return NextResponse.json({
      ok: true,
      monto,
      montoSoles,
      tipoCambio,
      presupuestoActualizado: true,
      celdaActualizada: `${hojaPresupuesto}!${direccionReal}`,
      gastoRealAnterior: valorActual,
      gastoRealNuevo: nuevoValor,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
