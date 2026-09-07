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
  ENCABEZADOS_NUEVOS_FACTURAS,
  extraerProyectos,
  fechaAExcelSerial,
  leerCeldaCruda,
  leerWorkbook,
  ultimaFilaConDatosEscaneada,
} from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/opex-constantes";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

interface CuerpoRegistro {
  filaProyecto: number;
  mes: number; // 1-12
  /** En qué moneda ingresó la persona `monto` — "PEN" (sin IGV, se convierte acá a USD
   *  con el tipo de cambio fijo de la app) o "USD" (ya viene en dólares, se usa tal cual). */
  moneda: "PEN" | "USD";
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
  if (moneda !== "PEN" && moneda !== "USD") {
    return NextResponse.json({ error: "Moneda inválida." }, { status: 400 });
  }
  if (!Number.isFinite(montoIngresado) || montoIngresado <= 0) {
    return NextResponse.json(
      { error: moneda === "PEN" ? "El monto en Soles debe ser mayor a 0." : "El monto en dólares debe ser mayor a 0." },
      { status: 400 }
    );
  }
  const fecha = new Date(periodoFacturado);
  if (Number.isNaN(fecha.getTime())) {
    return NextResponse.json({ error: "Periodo facturado inválido." }, { status: 400 });
  }

  const tipoCambio = TIPO_CAMBIO_POR_DEFECTO;
  // El USD es siempre lo que de verdad mueve el Gasto Real. Si ya se ingresó en dólares,
  // se usa tal cual; "Monto Soles (sin IGV)" solo se guarda cuando la factura de verdad
  // se originó en Soles — mismo criterio que ya usa OPEX.
  const monto = moneda === "USD" ? Math.round(montoIngresado * 100) / 100 : Math.round((montoIngresado / tipoCambio) * 100) / 100;
  const montoSoles = moneda === "PEN" ? montoIngresado : null;

  try {
    const archivo = await resolverArchivoPorShareUrl(config);
    const hojaProyectos = process.env.SP_CAPEX_HOJA?.trim() || "BD_CAPEX";

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
    //    corre ninguna columna existente.
    if (!leerCeldaCruda(wb, HOJA_FACTURAS, "J1")) {
      await escribirFila(config, archivo, HOJA_FACTURAS, 1, "J", "N", ENCABEZADOS_NUEVOS_FACTURAS);
    }

    // 3) Agrega la factura al final de la hoja de facturas. El mes al que pertenece el
    //    gasto ya no se embebe como texto en Comentarios ("Periodo X") — vive directo en
    //    su propia columna (Mes Real), así Comentarios queda libre para notas reales.
    const ultimaFila = ultimaFilaConDatosEscaneada(wb, HOJA_FACTURAS);
    const filaNueva = ultimaFila + 1;
    await escribirFila(config, archivo, HOJA_FACTURAS, filaNueva, "A", "N", [
      fechaAExcelSerial(fecha),
      recurso,
      proveedor,
      responsable,
      textoProyecto,
      monto,
      numeroFactura,
      "ok",
      comentarioExtra?.trim() ?? "",
      moneda,
      montoSoles ?? "",
      tipoCambio,
      ruc?.trim() ?? "",
      mes,
    ]);

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
      celdaActualizada: `${hojaProyectos}!${direccionReal}`,
      gastoRealAnterior: valorActual,
      gastoRealNuevo: nuevoValor,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}