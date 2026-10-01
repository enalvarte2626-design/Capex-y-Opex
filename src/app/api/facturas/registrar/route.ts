import { NextResponse } from "next/server";
import {
  ErrorSharePoint,
  camposFaltantes,
  descargarContenido,
  escribirCelda,
  escribirFila,
  leerCelda,
  obtenerConfiguracionSharePoint,
  resolverArchivoPorShareUrl,
} from "@/lib/sharepoint";
import {
  COL_BD,
  ENCABEZADOS_MONTO_EUROS,
  ENCABEZADOS_NUEVOS_FACTURAS,
  extraerProyectos,
  fechaAExcelSerial,
  leerCeldaCruda,
  leerWorkbook,
  ultimaFilaConDatosEscaneada,
} from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { TIPO_CAMBIO_EUR_POR_DEFECTO, TIPO_CAMBIO_POR_DEFECTO } from "@/lib/opex-constantes";
import { leerMesCierre, mesesCerradosPorNombreArchivo } from "@/lib/mesCierreConfig";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

interface CuerpoRegistro {
  filaProyecto: number;
  mes: number; // 1-12
  /** En qué moneda ingresó la persona `monto` — "PEN" (sin IGV, se convierte acá a USD
   *  con el tipo de cambio Soles↔Dólar de la app), "USD" (ya viene en dólares, se usa
   *  tal cual) o "EUR" (se convierte con el tipo de cambio Euro↔Dólar de la app). */
  moneda: "PEN" | "USD" | "EUR";
  /** El valor tal cual lo escribió la persona, en la moneda indicada por `moneda`. */
  monto: number;
  recurso: string;
  proveedor: string;
  responsable: string;
  numeroFactura: string;
  periodoFacturado: string; // "aaaa-mm-dd"
  comentarioExtra?: string;
  /** RUC del proveedor — opcional, solo aplica a proveedores peruanos. */
  ruc?: string;
}

export async function POST(request: Request) {
  const config = obtenerConfiguracionSharePoint();
  const faltantes = camposFaltantes(config);
  if (faltantes.length > 0) {
    return NextResponse.json(
      { error: `Falta configurar en .env.local: ${faltantes.join(", ")}.` },
      { status: 500 }
    );
  }

  const cuerpo = (await request.json().catch(() => null)) as CuerpoRegistro | null;
  if (!cuerpo) return NextResponse.json({ error: "Cuerpo inválido." }, { status: 400 });

  const {
    filaProyecto,
    mes,
    moneda,
    monto: montoIngresado,
    recurso,
    proveedor,
    responsable,
    numeroFactura,
    periodoFacturado,
    comentarioExtra,
    ruc,
  } = cuerpo;

  if (!Number.isInteger(filaProyecto) || filaProyecto < 2) {
    return NextResponse.json({ error: "Proyecto inválido." }, { status: 400 });
  }
  if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
    return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
  }
  if (moneda !== "PEN" && moneda !== "USD" && moneda !== "EUR") {
    return NextResponse.json({ error: "Moneda inválida." }, { status: 400 });
  }
  // Negativo se permite a propósito: es como se registra un descuento o nota de crédito
  // (resta del Gasto Real en vez de sumar, más abajo). Solo 0 no tiene sentido.
  if (!Number.isFinite(montoIngresado) || montoIngresado === 0) {
    const nombreMoneda = moneda === "PEN" ? "en Soles" : moneda === "EUR" ? "en Euros" : "en dólares";
    return NextResponse.json(
      { error: `El monto ${nombreMoneda} no puede ser 0 (usa negativo para un descuento o nota de crédito).` },
      { status: 400 }
    );
  }
  const fecha = new Date(periodoFacturado);
  if (Number.isNaN(fecha.getTime())) {
    return NextResponse.json({ error: "Periodo facturado inválido." }, { status: 400 });
  }

  const tipoCambio = TIPO_CAMBIO_POR_DEFECTO; // Soles↔Dólar — aplica siempre, sin importar la moneda ingresada
  const tipoCambioEur = TIPO_CAMBIO_EUR_POR_DEFECTO; // Euro↔Dólar — solo aplica si moneda === "EUR"
  // El USD es siempre lo que de verdad mueve el Gasto Real. Si ya se ingresó en dólares,
  // se usa tal cual; en Soles se divide por el tipo de cambio Soles↔Dólar; en Euros se
  // multiplica por el tipo de cambio Euro↔Dólar (1 EUR vale más que 1 USD). "Monto
  // Soles"/"Monto Euros" solo se guardan cuando la factura de verdad se originó en esa
  // moneda — mismo criterio que ya usa OPEX para Soles.
  const monto =
    moneda === "USD"
      ? Math.round(montoIngresado * 100) / 100
      : moneda === "EUR"
        ? Math.round(montoIngresado * tipoCambioEur * 100) / 100
        : Math.round((montoIngresado / tipoCambio) * 100) / 100;
  const montoSoles = moneda === "PEN" ? montoIngresado : null;
  const montoEuros = moneda === "EUR" ? montoIngresado : null;

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";

    // Un mes ya CERRADO sí se puede registrar — queda en el historial de facturas —
    // pero NUNCA suma al Gasto Real de BD_CAPEX, para no mover un presupuesto que ya se
    // presentó como cerrado. El "mes de cierre" es un marcador independiente (hoja
    // "Config App", mismo criterio que OPEX) que solo avanza desde el botón "Cerrar
    // mes" del Dashboard — generar el archivo del siguiente cierre (ej. "8+4" → "9+3")
    // YA NO cierra nada por sí solo, son dos decisiones separadas.
    const mesCierre = await leerMesCierre(config, archivo, Math.max(0, mesesCerradosPorNombreArchivo(archivo.nombre) - 1));
    const esMesPasado = mes <= mesCierre;

    // 1) Confirma que la fila de proyecto existe y arma el texto a guardar en "Proyecto".
    const contenido = await descargarContenido(config, archivo);
    const wb = leerWorkbook(contenido);
    const proyectos = extraerProyectos(wb, hojaProyectos);
    const proyecto = proyectos.find((p) => p.filaExcel === filaProyecto);
    if (!proyecto) {
      return NextResponse.json({ error: "No se encontró ese proyecto en BD_CAPEX." }, { status: 404 });
    }
    const textoProyecto = proyecto.detalle?.trim() || proyecto.proyecto;

    // 2) Si la hoja de facturas todavía no tiene las columnas J-N (Moneda/Monto
    //    Soles/Tipo de Cambio/RUC/Mes Real), agrega los encabezados una sola vez — nunca
    //    corre ninguna columna existente. O-P (Monto Euros/Tipo de Cambio Euro) se
    //    agregaron después, por eso se revisan aparte.
    if (!leerCeldaCruda(wb, HOJA_FACTURAS, "J1")) {
      await escribirFila(config, archivo, HOJA_FACTURAS, 1, "J", "N", ENCABEZADOS_NUEVOS_FACTURAS);
    }
    if (!leerCeldaCruda(wb, HOJA_FACTURAS, "O1")) {
      await escribirFila(config, archivo, HOJA_FACTURAS, 1, "O", "P", ENCABEZADOS_MONTO_EUROS);
    }

    // 3) Agrega la factura al final de la hoja de facturas. El mes al que pertenece el
    //    gasto ya no se embebe como texto en Comentarios ("Periodo X") — vive directo en
    //    su propia columna (Mes Real), así Comentarios queda libre para notas reales.
    const ultimaFila = ultimaFilaConDatosEscaneada(wb, HOJA_FACTURAS);
    const filaNueva = ultimaFila + 1;
    await escribirFila(config, archivo, HOJA_FACTURAS, filaNueva, "A", "P", [
      fechaAExcelSerial(fecha),
      recurso,
      proveedor,
      responsable,
      textoProyecto,
      monto,
      numeroFactura,
      esMesPasado ? "ok (mes pasado, no afecta presupuesto)" : "ok",
      comentarioExtra?.trim() ?? "",
      moneda,
      montoSoles ?? "",
      tipoCambio,
      ruc?.trim() ?? "",
      mes,
      montoEuros ?? "",
      moneda === "EUR" ? tipoCambioEur : "",
    ]);

    // Un mes pasado queda solo en el historial de facturas — nunca toca BD_CAPEX.
    if (esMesPasado) {
      return NextResponse.json({
        ok: true,
        filaFactura: filaNueva,
        monto,
        montoSoles,
        tipoCambio,
        montoEuros,
        tipoCambioEur: moneda === "EUR" ? tipoCambioEur : null,
        presupuestoActualizado: false,
        aviso: `Mes pasado: la factura quedó registrada en el historial, pero no se sumó al Gasto Real de ${hojaProyectos}.`,
      });
    }

    // 4) Suma el monto (USD) al Gasto Real del mes correspondiente en BD_CAPEX (no
    //    reemplaza: un proyecto puede tener varias facturas en el mismo mes).
    const colReal = columnaALetra(COL_BD.primerMesReal + (mes - 1) * 2);
    const direccionReal = `${colReal}${filaProyecto}`;
    const valorActual = await leerCelda(config, archivo, hojaProyectos, direccionReal);
    const nuevoValor = valorActual + monto;
    await escribirCelda(config, archivo, hojaProyectos, direccionReal, nuevoValor);

    return NextResponse.json({
      ok: true,
      filaFactura: filaNueva,
      monto,
      montoSoles,
      tipoCambio,
      montoEuros,
      tipoCambioEur: moneda === "EUR" ? tipoCambioEur : null,
      presupuestoActualizado: true,
      celdaActualizada: `${hojaProyectos}!${direccionReal}`,
      gastoRealAnterior: valorActual,
      gastoRealNuevo: nuevoValor,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}