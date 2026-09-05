import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantesOpex,
  crearHojaSiNoExiste,
  descargarContenido,
  escribirCelda,
  escribirFila,
  leerCelda,
  obtenerConfiguracionOpex,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import { COL_PPTO_OPEX, ENCABEZADOS_FACTURAS_OPEX, extraerPresupuestoOpex } from "@/lib/opex-parse";
import { fechaAExcelSerial, leerWorkbook, ultimaFilaConDatosEscaneada } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/opex-constantes";

export const dynamic = "force-dynamic";

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
  // No se permite registrar en un mes que ya pasó — evita sumar por error al Gasto Real
  // de un mes ya cerrado. Nunca se confía solo en que el formulario oculte esa opción.
  const mesActualReal = new Date().getMonth() + 1;
  if (mes < mesActualReal) {
    return NextResponse.json(
      { error: "No se puede registrar una factura en un mes que ya pasó. Elige el mes actual o uno futuro." },
      { status: 400 }
    );
  }
  if (!Number.isFinite(montoSoles) || montoSoles <= 0) {
    return NextResponse.json({ error: "El monto en Soles debe ser mayor a 0." }, { status: 400 });
  }
  if (!proveedor?.trim()) {
    return NextResponse.json({ error: "Falta el Proveedor." }, { status: 400 });
  }

  // Única fuente de verdad para el tipo de cambio de registro: nunca se recibe del
  // navegador, así ninguna factura puede quedar con una conversión manipulada.
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

    // 1) Confirma que la línea de gasto existe.
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const lineas = extraerPresupuestoOpex(wb, hojaPresupuesto);
    const linea = lineas.find((l) => l.filaExcel === filaPresupuesto);
    if (!linea) {
      return NextResponse.json({ error: "No se encontró esa línea de gasto en Presupuesto 2026." }, { status: 404 });
    }

    // 2) Crea la hoja de facturas (con encabezados) si es la primera vez que se usa.
    await crearHojaSiNoExiste(config, archivo, hojaFacturas);
    const contenidoActualizado = await descargarContenido(config, archivo);
    const wbActualizado = leerWorkbook(contenidoActualizado);
    const hojaExisteConDatos = wbActualizado.Sheets[hojaFacturas]?.["!ref"] != null;
    if (!hojaExisteConDatos) {
      // Hoja recién creada, todavía sin nada — escribe el encabezado.
      await escribirFila(config, archivo, hojaFacturas, 1, "A", "O", ENCABEZADOS_FACTURAS_OPEX);
    }

    // 3) Agrega la factura al final de esa hoja.
    const ultimaFila = ultimaFilaConDatosEscaneada(wbActualizado, hojaFacturas);
    const filaNueva = Math.max(ultimaFila, 1) + 1;
    await escribirFila(config, archivo, hojaFacturas, filaNueva, "A", "O", [
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
      "ok",
      montoSoles,
      tipoCambio,
      linea.subNegocio,
    ]);

    // 4) Suma el monto (USD) al Gasto Real del mes correspondiente en Presupuesto 2026.
    const colReal = columnaALetra(COL_PPTO_OPEX.primerMesReal + (mes - 1) * 2);
    const direccionReal = `${colReal}${filaPresupuesto}`;
    const valorActual = await leerCelda(config, archivo, hojaPresupuesto, direccionReal);
    const nuevoValor = valorActual + monto;
    await escribirCelda(config, archivo, hojaPresupuesto, direccionReal, nuevoValor);

    return NextResponse.json({
      ok: true,
      filaFactura: filaNueva,
      monto,
      montoSoles,
      tipoCambio,
      celdaActualizada: `${hojaPresupuesto}!${direccionReal}`,
      gastoRealAnterior: valorActual,
      gastoRealNuevo: nuevoValor,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}
