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
import { COL_PPTO_OPEX, ENCABEZADOS_FACTURAS_OPEX, extraerPresupuestoOpex } from "@/lib/opex-parse";
import { fechaAExcelSerial, leerWorkbook } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/opex-constantes";

export const dynamic = "force-dynamic";

interface CuerpoRegistro {
  filaPresupuesto: number;
  mes: number; // 1-12
  /** En qué moneda ingresó la persona `monto` — "PEN" (Soles sin IGV, se convierte acá
   *  a USD) o "USD" (ya viene en dólares, se usa tal cual). */
  moneda: "PEN" | "USD";
  /** El valor tal cual lo escribió la persona, en la moneda indicada por `moneda`. */
  monto: number;
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
 * El monto se puede ingresar en Soles SIN IGV (así llegan la mayoría de las facturas
 * locales) o directo en Dólares (para proveedores extranjeros que ya facturan en USD).
 * Cuando es en Soles, se convierte acá a USD con el tipo de cambio fijo de la app
 * (`TIPO_CAMBIO_POR_DEFECTO`, hoy 3.4) — ese USD es el que de verdad se suma al Gasto
 * Real; nunca se confía en un monto en USD calculado del lado del navegador.
 */
export async function POST(request: Request) {
  const config = obtenerConfiguracionOpex();
  const faltantes = camposFaltantesOpex(config);
  if (faltantes.length > 0) {
    return NextResponse.json({ error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` }, { status: 500 });
  }

  const cuerpo = (await request.json().catch(() => null)) as CuerpoRegistro | null;
  if (!cuerpo) return NextResponse.json({ error: "Cuerpo inválido." }, { status: 400 });

  const { filaPresupuesto, mes, moneda, monto: montoIngresado, proveedor, numeroComprobante, comentario } = cuerpo;

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
  if (moneda !== "PEN" && moneda !== "USD") {
    return NextResponse.json({ error: "Moneda inválida." }, { status: 400 });
  }
  if (!Number.isFinite(montoIngresado) || montoIngresado <= 0) {
    return NextResponse.json(
      { error: moneda === "PEN" ? "El monto en Soles debe ser mayor a 0." : "El monto en dólares debe ser mayor a 0." },
      { status: 400 }
    );
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
  // El USD es siempre lo que de verdad mueve el Gasto Real. Si la persona ya ingresó en
  // dólares, se usa tal cual (sin dividir de nuevo por el tipo de cambio); el "Monto
  // Soles (sin IGV)" solo se guarda cuando la factura de verdad se originó en Soles.
  const monto = moneda === "USD" ? Math.round(montoIngresado * 100) / 100 : Math.round((montoIngresado / tipoCambio) * 100) / 100;
  const montoSoles = moneda === "PEN" ? montoIngresado : null;

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
      await escribirFila(config, archivo, hojaFacturas, 1, "A", "P", ENCABEZADOS_FACTURAS_OPEX);
    }

    // 3) Agrega la factura al final de esa hoja — vía la Tabla de Excel de la hoja (no
    // calculando nosotros "última fila + 1"), para no pisar otra factura que se haya
    // registrado casi al mismo tiempo. Ver el comentario de `agregarFilaTabla` en
    // sharepoint.ts.
    await asegurarTablaEnHoja(config, archivo, hojaFacturas, "A1:P1");
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
      montoSoles ?? "",
      tipoCambio,
      linea.subNegocio,
      moneda,
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

    // 4) Suma el monto (USD) al Gasto Real del mes correspondiente en Presupuesto 2026.
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
