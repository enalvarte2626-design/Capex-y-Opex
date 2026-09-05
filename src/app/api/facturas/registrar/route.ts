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
import { COL_BD, extraerProyectos, fechaAExcelSerial, leerWorkbook, ultimaFilaConDatosEscaneada } from "@/lib/capex-parse";
import { columnaALetra } from "@/lib/capex-editable";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";

export const dynamic = "force-dynamic";

const HOJA_FACTURAS = "Control de Facturas-Capex 25fEB";

interface CuerpoRegistro {
  filaProyecto: number;
  mes: number; // 1-12
  monto: number;
  recurso: string;
  proveedor: string;
  responsable: string;
  numeroFactura: string;
  periodoFacturado: string; // "aaaa-mm-dd"
  comentarioExtra?: string;
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

  const { filaProyecto, mes, monto, recurso, proveedor, responsable, numeroFactura, periodoFacturado, comentarioExtra } =
    cuerpo;

  if (!Number.isInteger(filaProyecto) || filaProyecto < 2) {
    return NextResponse.json({ error: "Proyecto inválido." }, { status: 400 });
  }
  if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
    return NextResponse.json({ error: "Mes inválido." }, { status: 400 });
  }
  if (!Number.isFinite(monto) || monto <= 0) {
    return NextResponse.json({ error: "El monto debe ser mayor a 0." }, { status: 400 });
  }
  const fecha = new Date(periodoFacturado);
  if (Number.isNaN(fecha.getTime())) {
    return NextResponse.json({ error: "Periodo facturado inválido." }, { status: 400 });
  }

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

    // 2) Agrega la factura al final de la hoja de facturas.
    const ultimaFila = ultimaFilaConDatosEscaneada(wb, HOJA_FACTURAS);
    const filaNueva = ultimaFila + 1;
    const comentario = `Periodo ${NOMBRES_MES_CIERRE[mes - 1]}${comentarioExtra ? ` — ${comentarioExtra}` : ""}`;
    await escribirFila(config, archivo, HOJA_FACTURAS, filaNueva, "A", "I", [
      fechaAExcelSerial(fecha),
      recurso,
      proveedor,
      responsable,
      textoProyecto,
      monto,
      numeroFactura,
      "ok",
      comentario,
    ]);

    // 3) Suma el monto al Gasto Real del mes correspondiente en BD_CAPEX (no reemplaza:
    //    un proyecto puede tener varias facturas en el mismo mes).
    const colReal = columnaALetra(COL_BD.primerMesReal + (mes - 1) * 2);
    const direccionReal = `${colReal}${filaProyecto}`;
    const valorActual = await leerCelda(config, archivo, hojaProyectos, direccionReal);
    const nuevoValor = valorActual + monto;
    await escribirCelda(config, archivo, hojaProyectos, direccionReal, nuevoValor);

    return NextResponse.json({
      ok: true,
      filaFactura: filaNueva,
      celdaActualizada: `${hojaProyectos}!${direccionReal}`,
      gastoRealAnterior: valorActual,
      gastoRealNuevo: nuevoValor,
    });
  } catch (e) {
    const mensaje = e instanceof ErrorSharePoint ? e.message : `Error inesperado: ${(e as Error).message}`;
    return NextResponse.json({ error: mensaje }, { status: 502 });
  }
}