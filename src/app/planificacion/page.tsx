"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE, resolverProyectos, type ProyectoCapex } from "@/lib/capex";
import { moneda2, soles } from "@/lib/format";
import { useMesCierre } from "@/lib/useMesCierre";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import { useNivelAcceso } from "@/lib/useNivelAcceso";
import CampoMontoSumado from "@/components/CampoMontoSumado";
import ControlTipoCambio from "@/components/ControlTipoCambio";
import { useAnio } from "@/components/AnioProvider";

const NOMBRES_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const RANGO_DIACRITICOS = /[̀-ͯ]/g;

function normalizar(texto: string): string {
  return texto.toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
}

/** Umbral para considerar una línea "programada": menos de medio centavo de diferencia
 *  entre el Presupuesto Aprobado y lo ya repartido (Real + Forecast) entre los 12 meses. */
const TOLERANCIA_DIFERENCIA = 0.005;

const ESTILO_DESTACADO = { background: "var(--acento-suave)" };

/**
 * Planificación CAPEX 2026: a diferencia de "Detalle BD_CAPEX" (que mezcla avance, status,
 * facturas y meses en una sola tabla enorme), esta pantalla está pensada para una sola
 * tarea — repartir el Presupuesto Aprobado de cada proyecto entre los 12 meses del año —
 * así distintas personas pueden entrar solo a programar, sin el resto del ruido. La
 * columna "Diferencia" (Presupuesto Aprobado − Gasto Real − Forecast, igual que en el
 * Dashboard) es la que dice si ya quedó todo repartido: en $0 significa que el año está
 * completamente programado; distinto de $0 significa que falta (o sobra) plata por
 * asignar a algún mes.
 */
export default function PlanificacionCapex() {
  const nivelAcceso = useNivelAcceso();
  const puedeEditar = nivelAcceso === "completo";
  const anio = useAnio();

  const [proyectos, setProyectos] = useState<ProyectoCapex[] | null>(null);
  const [archivo, setArchivo] = useState("");
  const [actualizadoEn, setActualizadoEn] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [busqueda, setBusqueda] = useState("");
  const [grupoSel, setGrupoSel] = usePersistedState<string | null>("planificacion-capex-grupo", null);
  const [soloSinProgramar, setSoloSinProgramar] = usePersistedState("planificacion-capex-sin-programar", false);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("planificacion-capex-soles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();
  const [mesCierre, setMesCierre] = useMesCierre();

  function actualizarLocal(fila: number, cambios: Partial<ProyectoCapex>) {
    setProyectos((prev) => prev?.map((x) => (x.filaExcel === fila ? { ...x, ...cambios } : x)) ?? prev);
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/capex", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setProyectos(json.proyectos);
      setArchivo(json.archivo);
      setActualizadoEn(json.actualizadoEn);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
  }, []);

  const resueltos = useMemo(() => resolverProyectos(proyectos ?? [], mesCierre), [proyectos, mesCierre]);

  const meses = useMemo(
    () => NOMBRES_MES.map((nombre, mi) => ({ nombre, mi, cerrado: mi < mesCierre })),
    [mesCierre]
  );

  const grupos = useMemo(
    () => Array.from(new Set(resueltos.map((p) => p.grupoNegocio || "SIN GRUPO"))).sort((a, b) => a.localeCompare(b, "es")),
    [resueltos]
  );

  const sinProgramarTotal = useMemo(
    () => resueltos.filter((p) => Math.abs(p.diferencia) > TOLERANCIA_DIFERENCIA).length,
    [resueltos]
  );

  const filtrados = useMemo(() => {
    let filas = resueltos;
    if (grupoSel) filas = filas.filter((p) => (p.grupoNegocio || "SIN GRUPO") === grupoSel);
    if (soloSinProgramar) filas = filas.filter((p) => Math.abs(p.diferencia) > TOLERANCIA_DIFERENCIA);
    const q = normalizar(busqueda);
    if (q) filas = filas.filter((p) => normalizar(p.proyecto).includes(q) || normalizar(p.detalle).includes(q));
    return [...filas].sort(
      (a, b) => a.grupoNegocio.localeCompare(b.grupoNegocio, "es") || a.proyecto.localeCompare(b.proyecto, "es")
    );
  }, [resueltos, grupoSel, soloSinProgramar, busqueda]);

  const totales = useMemo(
    () =>
      filtrados.reduce(
        (acc, p) => ({
          presupuestoAprobado: acc.presupuestoAprobado + p.presupuestoAprobado,
          gastoReal: acc.gastoReal + p.gastoReal,
          forecast: acc.forecast + p.forecast,
          diferencia: acc.diferencia + p.diferencia,
        }),
        { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 }
      ),
    [filtrados]
  );

  if (cargando && !proyectos) {
    return <p style={{ color: "var(--texto-suave)" }}>Cargando datos desde SharePoint…</p>;
  }

  if (error) {
    return (
      <div className="card p-6">
        <p className="font-semibold mb-1" style={{ color: "var(--peligro)" }}>
          No se pudo cargar el archivo
        </p>
        <p className="text-sm mb-4" style={{ color: "var(--texto-suave)" }}>
          {error}
        </p>
        <button className="boton-primario" onClick={cargar}>
          Reintentar
        </button>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4">
      <div>
        <h2 className="text-lg font-semibold">Planificación CAPEX {anio}</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Reparte el Presupuesto Aprobado de cada proyecto entre los 12 meses del año ({archivo || "…"}). La
          columna Diferencia muestra lo que todavía falta (o sobra) por programar — en $0 el proyecto ya quedó
          totalmente repartido. {filtrados.length} de {proyectos?.length ?? 0} proyectos.
        </p>
      </div>

      <div className="card p-4 flex flex-col gap-3">
        <div className="flex flex-wrap items-center gap-3">
          <input
            type="text"
            placeholder="Buscar proyecto o detalle…"
            className="campo"
            style={{ maxWidth: 260 }}
            value={busqueda}
            onChange={(e) => setBusqueda(e.target.value)}
          />
          <select
            className="campo"
            style={{ width: "auto" }}
            value={grupoSel ?? ""}
            onChange={(e) => setGrupoSel(e.target.value || null)}
          >
            <option value="">Todos los grupos</option>
            {grupos.map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
          </select>
          <span
            className="chip"
            data-activo={soloSinProgramar}
            onClick={() => setSoloSinProgramar((v) => !v)}
            title="Muestra solo los proyectos donde el Presupuesto Aprobado todavía no está totalmente repartido en los 12 meses"
          >
            {soloSinProgramar ? "✓ " : ""}Solo sin programar{sinProgramarTotal > 0 ? ` (${sinProgramarTotal})` : ""}
          </span>
          <span
            className="chip"
            data-activo={mostrarSoles}
            onClick={() => setMostrarSoles((v) => !v)}
            title="Muestra el equivalente en Soles como referencia — no afecta ningún cálculo ni se guarda en el Excel."
          >
            {mostrarSoles ? "✓ " : ""}Habilitar en Soles
          </span>
          {mostrarSoles && <ControlTipoCambio tipoCambio={tipoCambio} onCambiar={setTipoCambio} />}
          <div className="flex items-center gap-2">
            <span className="etiqueta mb-0">Mes de cierre:</span>
            <select
              className="campo"
              style={{ width: "auto" }}
              value={mesCierre}
              onChange={(e) => setMesCierre(Number(e.target.value))}
              title="Último mes con Gasto Real ya cerrado — antes de ese mes no se programa, solo se ve lo ya ejecutado."
            >
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  {nombre}
                </option>
              ))}
            </select>
          </div>
          <div className="flex items-center gap-3 ml-auto">
            {actualizadoEn && (
              <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                Actualizado {new Date(actualizadoEn).toLocaleString("es-PE")}
              </span>
            )}
            <button className="boton-primario" onClick={cargar} disabled={cargando}>
              {cargando ? "Actualizando…" : "Actualizar"}
            </button>
          </div>
        </div>
        {!puedeEditar && (
          <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
            Estás con acceso de solo lectura: puedes ver la planificación, pero no puedes cambiar ningún mes.
          </p>
        )}
      </div>

      <div className="card p-0 overflow-hidden">
        <div style={{ maxHeight: "75vh", overflow: "auto" }}>
          <table className="text-xs border-collapse" style={{ tableLayout: "fixed", width: 1660 }}>
            <colgroup>
              <col style={{ width: 40 }} />
              <col style={{ width: 190 }} />
              <col style={{ width: 190 }} />
              <col style={{ width: 90 }} />
              <col style={{ width: 115 }} />
              {meses.map((m) => (
                <col key={m.mi} style={{ width: 90 }} />
              ))}
              <col style={{ width: 100 }} />
            </colgroup>
            <thead style={{ position: "sticky", top: 0, zIndex: 10, background: "var(--bg)" }}>
              <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--bg)" }}>
                <th className="py-2 px-3 font-semibold text-center">#</th>
                <th className="py-2 px-3 font-semibold">Proyecto</th>
                <th className="py-2 px-3 font-semibold">Detalle</th>
                <th className="py-2 px-3 font-semibold">Grupo</th>
                <th className="py-2 px-3 font-semibold text-center" style={ESTILO_DESTACADO}>
                  Presupuesto Aprobado
                </th>
                {meses.map((m) => (
                  <th key={m.mi} className="py-2 px-3 font-semibold text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                    {m.nombre}
                    {m.cerrado && (
                      <span className="block text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
                        cerrado
                      </span>
                    )}
                  </th>
                ))}
                <th className="py-2 px-3 font-semibold text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                  Diferencia
                </th>
              </tr>
            </thead>
            <tbody>
              {filtrados.map((p, i) => {
                const sinProgramar = Math.abs(p.diferencia) > TOLERANCIA_DIFERENCIA;
                return (
                  <tr
                    key={p.filaExcel}
                    style={{ borderTop: "1px solid var(--borde)", background: sinProgramar ? "#fff8e6" : undefined }}
                  >
                    <td className="py-1.5 px-3 text-center" style={{ color: "var(--texto-suave)" }}>
                      {i + 1}
                    </td>
                    <td className="py-1.5 px-3 truncate" title={p.proyecto}>
                      {p.proyecto || "—"}
                    </td>
                    <td className="py-1.5 px-3 truncate" style={{ color: "var(--texto-suave)" }} title={p.detalle}>
                      {p.detalle || "—"}
                    </td>
                    <td className="py-1.5 px-3 truncate">{p.grupoNegocio}</td>
                    <td className="py-1.5 px-3 text-center whitespace-nowrap font-medium" style={ESTILO_DESTACADO}>
                      {moneda2(p.presupuestoAprobado)}
                      {mostrarSoles && (
                        <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                          {soles(p.presupuestoAprobado, tipoCambio)}
                        </span>
                      )}
                    </td>
                    {meses.map((m) => (
                      <td key={m.mi} className="py-1.5 px-3" style={{ borderLeft: "1px solid var(--borde)" }}>
                        {m.cerrado ? (
                          <span
                            className="block text-center text-xs"
                            style={{ color: "var(--texto-suave)" }}
                            title="Mes cerrado — ya se ejecutó, no se programa"
                          >
                            {moneda2(p.real[m.mi])}
                          </span>
                        ) : (
                          <CampoMontoSumado
                            fila={p.filaExcel}
                            campo={`proyectado:${m.mi}`}
                            valor={p.proyectado[m.mi]}
                            className="text-center text-xs"
                            soloLectura={!puedeEditar}
                            mostrarSoles={mostrarSoles}
                            tipoCambio={tipoCambio}
                            onGuardado={(nuevo) =>
                              actualizarLocal(p.filaExcel, { proyectado: p.proyectado.map((x, j) => (j === m.mi ? nuevo : x)) })
                            }
                          />
                        )}
                      </td>
                    ))}
                    <td
                      className="py-1.5 px-3 text-center whitespace-nowrap font-bold"
                      style={{
                        borderLeft: "1px solid var(--borde)",
                        color: p.diferencia < -TOLERANCIA_DIFERENCIA
                          ? "var(--peligro)"
                          : p.diferencia > TOLERANCIA_DIFERENCIA
                            ? "var(--alerta)"
                            : "var(--exito)",
                      }}
                      title={
                        p.diferencia < -TOLERANCIA_DIFERENCIA
                          ? "Ya se programó (o gastó) más de lo aprobado"
                          : p.diferencia > TOLERANCIA_DIFERENCIA
                            ? "Todavía queda presupuesto sin repartir en algún mes"
                            : "Presupuesto totalmente programado"
                      }
                    >
                      {moneda2(p.diferencia)}
                    </td>
                  </tr>
                );
              })}
              {filtrados.length === 0 && (
                <tr>
                  <td colSpan={6 + meses.length} className="py-6 text-center" style={{ color: "var(--texto-suave)" }}>
                    Sin proyectos para este filtro.
                  </td>
                </tr>
              )}
            </tbody>
            <tfoot>
              <tr style={{ borderTop: "2px solid var(--borde)" }}>
                <td className="py-2 px-3 font-bold" colSpan={4} style={{ background: "var(--bg)" }}>
                  Total ({filtrados.length})
                </td>
                <td className="py-2 px-3 text-center font-bold" style={{ background: "var(--bg)" }}>
                  {moneda2(totales.presupuestoAprobado)}
                </td>
                <td colSpan={meses.length} style={{ background: "var(--bg)" }}></td>
                <td
                  className="py-2 px-3 text-center font-bold"
                  style={{
                    background: "var(--bg)",
                    borderLeft: "1px solid var(--borde)",
                    color: totales.diferencia < -TOLERANCIA_DIFERENCIA ? "var(--peligro)" : "var(--exito)",
                  }}
                >
                  {moneda2(totales.diferencia)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}
