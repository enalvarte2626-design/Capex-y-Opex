"use client";

import { useEffect, useState } from "react";
import { moneda2 } from "@/lib/format";

const ROSA_TEMPLATE = "#d6246e";

interface ArchivoCierreDisponible {
  nombre: string;
  cerrados: number;
}

interface CambioLinea {
  filaExcel: number;
  proyecto: string;
  detalle: string;
  grupoNegocio: string;
  campo: string;
  valorAnterior: number;
  valorActual: number;
  diferencia: number;
  eraMesCerrado: boolean;
}

interface ResultadoComparacion {
  archivoAnterior: string;
  cerradosEnArchivoAnterior: number;
  proyectosNuevos: string[];
  cambios: CambioLinea[];
}

/**
 * Compara BD_CAPEX en vivo contra un archivo de cierre anterior (los que va dejando
 * "Generar archivo de cierre" al lado, sin tocarlos) — para responder "¿qué se movió
 * desde el cierre pasado que hace que un indicador salga distinto ahora?". Los cambios
 * en un mes que YA estaba cerrado en ese archivo anterior se muestran primero y
 * resaltados: son la causa más probable de un indicador cerrado que "se mueve solo".
 */
export default function CompararCierre() {
  const [archivos, setArchivos] = useState<ArchivoCierreDisponible[] | null>(null);
  const [archivoActual, setArchivoActual] = useState("");
  const [seleccionado, setSeleccionado] = useState("");
  const [resultado, setResultado] = useState<ResultadoComparacion | null>(null);
  const [cargandoLista, setCargandoLista] = useState(true);
  const [comparando, setComparando] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [soloMesesCerrados, setSoloMesesCerrados] = useState(false);

  useEffect(() => {
    fetch("/api/capex/cierres-disponibles", { cache: "no-store" })
      .then((r) => r.json())
      .then((j) => {
        if (j.error) throw new Error(j.error);
        setArchivos(j.disponibles);
        setArchivoActual(j.archivoActual);
        if (j.disponibles.length > 0) setSeleccionado(j.disponibles[0].nombre);
      })
      .catch((e) => setError((e as Error).message))
      .finally(() => setCargandoLista(false));
  }, []);

  async function comparar() {
    if (!seleccionado) return;
    setComparando(true);
    setError(null);
    setResultado(null);
    try {
      const res = await fetch(`/api/capex/comparar-cierre?archivo=${encodeURIComponent(seleccionado)}`, { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo comparar.");
      setResultado(json);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setComparando(false);
    }
  }

  const cambiosMostrados = resultado
    ? soloMesesCerrados
      ? resultado.cambios.filter((c) => c.eraMesCerrado)
      : resultado.cambios
    : [];
  const totalMesesCerrados = resultado?.cambios.filter((c) => c.eraMesCerrado).length ?? 0;

  return (
    <div className="flex flex-col gap-4">
      <div>
        <h2 className="text-lg font-semibold">Comparar con un cierre anterior</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Compara BD_CAPEX en vivo ({archivoActual || "…"}) contra el archivo que quedó guardado la última vez
          que se usó &quot;Generar archivo de cierre&quot; (o uno más antiguo, si lo eliges) — para ver exactamente
          qué celda cambió desde entonces, mes por mes.
        </p>
      </div>

      {cargandoLista && <p style={{ color: "var(--texto-suave)" }}>Buscando archivos de cierre anteriores…</p>}

      {!cargandoLista && archivos && archivos.length === 0 && (
        <div className="card p-6 text-center" style={{ color: "var(--texto-suave)" }}>
          No hay ningún archivo de cierre anterior en la misma carpeta — se generan con el botón
          &quot;Generar archivo de cierre&quot; en el Dashboard, cada vez que se cierra un mes.
        </div>
      )}

      {!cargandoLista && archivos && archivos.length > 0 && (
        <div className="card p-4 flex flex-wrap items-center gap-3">
          <span className="etiqueta mb-0">Comparar contra:</span>
          <select className="campo" style={{ width: "auto", minWidth: 280 }} value={seleccionado} onChange={(e) => setSeleccionado(e.target.value)}>
            {archivos.map((a) => (
              <option key={a.nombre} value={a.nombre}>
                {a.nombre} ({a.cerrados} meses cerrados en ese momento)
              </option>
            ))}
          </select>
          <button className="boton-primario" onClick={comparar} disabled={comparando}>
            {comparando ? "Comparando…" : "Comparar"}
          </button>
        </div>
      )}

      {error && (
        <div className="card p-3 text-sm" style={{ background: "#fbe1ec", color: "var(--peligro)" }}>
          {error}
        </div>
      )}

      {resultado && (
        <>
          {resultado.proyectosNuevos.length > 0 && (
            <div className="card p-4 text-sm">
              <p className="font-semibold mb-1">{resultado.proyectosNuevos.length} proyecto(s) nuevo(s) desde entonces</p>
              <p style={{ color: "var(--texto-suave)" }}>
                No existían todavía en {resultado.archivoAnterior}, así que no hay con qué compararlos:{" "}
                {resultado.proyectosNuevos.join(", ")}.
              </p>
            </div>
          )}

          {resultado.cambios.length === 0 ? (
            <div className="card p-6 text-center" style={{ color: "var(--exito)" }}>
              Ningún proyecto cambió (Presupuesto Aprobado, Real ni Proyectado) desde {resultado.archivoAnterior}.
            </div>
          ) : (
            <div className="card p-0 overflow-hidden">
              <div className="p-4 flex items-center gap-3" style={{ borderBottom: "1px solid var(--borde)" }}>
                <span
                  className="chip"
                  data-activo={soloMesesCerrados}
                  onClick={() => setSoloMesesCerrados((v) => !v)}
                  title="Muestra solo cambios en un mes que ya estaba cerrado en el archivo anterior — la causa más probable de un indicador que 'se mueve solo'"
                >
                  {soloMesesCerrados ? "✓ " : ""}Solo meses ya cerrados{totalMesesCerrados > 0 ? ` (${totalMesesCerrados})` : ""}
                </span>
                <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                  {cambiosMostrados.length} de {resultado.cambios.length} cambios
                </span>
              </div>
              <div style={{ maxHeight: "65vh", overflow: "auto" }}>
                <table className="text-xs border-collapse w-full">
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)" }}>
                    <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                      <th className="py-2 px-3 font-semibold">Proyecto</th>
                      <th className="py-2 px-3 font-semibold">Campo</th>
                      <th className="py-2 px-3 font-semibold text-center">Antes</th>
                      <th className="py-2 px-3 font-semibold text-center">Ahora</th>
                      <th className="py-2 px-3 font-semibold text-center">Diferencia</th>
                    </tr>
                  </thead>
                  <tbody>
                    {cambiosMostrados.map((c, i) => (
                      <tr
                        key={i}
                        style={{
                          borderTop: "1px solid var(--borde)",
                          background: c.eraMesCerrado ? "#fbe1ec" : undefined,
                        }}
                      >
                        <td className="py-1.5 px-3">
                          <span className="font-semibold">{c.proyecto || "—"}</span>
                          {c.detalle && (
                            <span className="block" style={{ color: "var(--texto-suave)" }}>
                              {c.detalle}
                            </span>
                          )}
                        </td>
                        <td className="py-1.5 px-3">
                          {c.campo}
                          {c.eraMesCerrado && (
                            <span className="block font-semibold" style={{ color: ROSA_TEMPLATE }}>
                              ⚠ mes ya cerrado entonces
                            </span>
                          )}
                        </td>
                        <td className="py-1.5 px-3 text-center whitespace-nowrap">{moneda2(c.valorAnterior)}</td>
                        <td className="py-1.5 px-3 text-center whitespace-nowrap">{moneda2(c.valorActual)}</td>
                        <td
                          className="py-1.5 px-3 text-center whitespace-nowrap font-bold"
                          style={{ color: c.diferencia < 0 ? "var(--peligro)" : "var(--exito)" }}
                        >
                          {c.diferencia > 0 ? "+" : ""}
                          {moneda2(c.diferencia)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}
