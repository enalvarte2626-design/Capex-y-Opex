"use client";

import { useEffect, useMemo, useState } from "react";
import { resolverProyectos, type ProyectoCapex } from "@/lib/capex";
import { moneda2, soles } from "@/lib/format";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import { useNivelAcceso } from "@/lib/useNivelAcceso";
import CampoEditable from "@/components/CampoEditable";
import CampoMontoSumado from "@/components/CampoMontoSumado";
import ControlTipoCambio from "@/components/ControlTipoCambio";
import { useAnios } from "@/components/AnioProvider";

const NOMBRES_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const RANGO_DIACRITICOS = /[̀-ͯ]/g;
const GRUPOS_NEGOCIO = ["EMISIVO", "RECEPTIVO", "TRANSVERSAL"];
const ENDPOINT_CELDA = "/api/capex/planificacion/celda";

function normalizar(texto: string): string {
  return texto.toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
}

/** Umbral para considerar una línea "programada": menos de medio centavo de diferencia
 *  entre el Presupuesto Aprobado y lo ya repartido entre los 12 meses. */
const TOLERANCIA_DIFERENCIA = 0.005;

const ESTILO_DESTACADO = { background: "var(--acento-suave)" };

/**
 * Planificación CAPEX: el borrador del presupuesto del año SIGUIENTE al activo — vive en
 * su propia hoja de Excel ("BD_CAPEX (Planificación)"), separada de BD_CAPEX (el año en
 * curso), así que editar acá nunca toca lo que ya se está ejecutando. El flujo es:
 * 1) "Generar borrador" copia la lista de proyectos vigente hoy, en blanco (Presupuesto
 *    Aprobado y meses en $0) — se puede repetir cuando se quiera empezar de nuevo.
 * 2) Se reparte el Presupuesto Aprobado de cada proyecto entre los 12 meses (o se agregan
 *    proyectos nuevos que todavía no existen).
 * 3) "Aprobar y activar" hace que este borrador PASE A SER el BD_CAPEX en vivo — de ahí
 *    en adelante todos los indicadores de la app salen de acá. El BD_CAPEX saliente queda
 *    archivado con el año en el nombre, nunca se pierde.
 */
export default function PlanificacionCapex() {
  const nivelAcceso = useNivelAcceso();
  const puedeEditar = nivelAcceso === "completo";
  const anioActivoApp = useAnios().capex;

  const [proyectos, setProyectos] = useState<ProyectoCapex[] | null>(null);
  const [anioBorrador, setAnioBorrador] = useState<number | null>(null);
  const [archivo, setArchivo] = useState("");
  const [actualizadoEn, setActualizadoEn] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [generando, setGenerando] = useState(false);
  const [aprobando, setAprobando] = useState(false);
  const [mensaje, setMensaje] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);
  const [mostrarFormNuevo, setMostrarFormNuevo] = useState(false);

  const [busqueda, setBusqueda] = useState("");
  const [grupoSel, setGrupoSel] = usePersistedState<string | null>("planificacion-capex-grupo", null);
  const [soloSinProgramar, setSoloSinProgramar] = usePersistedState("planificacion-capex-sin-programar", false);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("planificacion-capex-soles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();

  function actualizarLocal(fila: number, cambios: Partial<ProyectoCapex>) {
    setProyectos((prev) => prev?.map((x) => (x.filaExcel === fila ? { ...x, ...cambios } : x)) ?? prev);
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/capex/planificacion", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setProyectos(json.proyectos);
      setAnioBorrador(json.anioBorrador);
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

  async function generarBorrador() {
    if (
      (proyectos?.length ?? 0) > 0 &&
      !window.confirm(
        "Ya hay un borrador con datos. Generarlo de nuevo BORRA todo lo que se haya escrito acá (vuelve a copiar la lista de proyectos vigente, en blanco). ¿Continuar?"
      )
    ) {
      return;
    }
    setGenerando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/capex/planificacion/generar", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo generar el borrador.");
      setMensaje({ tipo: "ok", texto: `Borrador generado con ${json.lineas} proyectos.` });
      await cargar();
    } catch (e) {
      setMensaje({ tipo: "error", texto: (e as Error).message });
    } finally {
      setGenerando(false);
    }
  }

  async function aprobarYActivar() {
    if (
      !window.confirm(
        `¿Aprobar y activar ${anioBorrador}? Desde este momento BD_CAPEX (y todos los indicadores de la app) van a ser los de este borrador. El BD_CAPEX de ${anioActivoApp} queda archivado con el año en el nombre — no se pierde, pero deja de ser el que se ve en el Dashboard. Esta acción no se puede deshacer desde la app.`
      )
    ) {
      return;
    }
    setAprobando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/capex/planificacion/aprobar", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo aprobar y activar.");
      setMensaje({
        tipo: "ok",
        texto: `${json.anioNuevo - 1} archivado como "${json.hojaArchivada}". BD_CAPEX ${json.anioNuevo} ya está en vivo.`,
      });
      await cargar();
      // El año activo repartido por el layout (menú, títulos) queda desactualizado hasta
      // el próximo request al servidor — refresca toda la página para que se note ya.
      window.location.href = "/";
    } catch (e) {
      setMensaje({ tipo: "error", texto: (e as Error).message });
    } finally {
      setAprobando(false);
    }
  }

  // mesCierre=0: en un borrador todavía no hay ningún mes ejecutado — los 12 son
  // "Proyectado" (Forecast), así "diferencia" queda como Presupuesto Aprobado menos lo
  // ya repartido, exactamente lo que hace falta para saber si falta programar algo.
  const resueltos = useMemo(() => resolverProyectos(proyectos ?? [], 0), [proyectos]);

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
        (acc, p) => ({ presupuestoAprobado: acc.presupuestoAprobado + p.presupuestoAprobado, diferencia: acc.diferencia + p.diferencia }),
        { presupuestoAprobado: 0, diferencia: 0 }
      ),
    [filtrados]
  );

  if (cargando && proyectos === null) {
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
        <h2 className="text-lg font-semibold">Planificación CAPEX {anioBorrador ?? ""}</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Borrador del presupuesto {anioBorrador} — separado de BD_CAPEX ({anioActivoApp}, {archivo || "…"}), no
          afecta nada en vivo hasta que se apruebe. {filtrados.length} de {proyectos?.length ?? 0} proyectos.
        </p>
      </div>

      {mensaje && (
        <div
          className="card p-3 text-sm"
          style={{
            background: mensaje.tipo === "ok" ? "#e3f3e3" : "#fbe1ec",
            color: mensaje.tipo === "ok" ? "var(--exito)" : "var(--peligro)",
          }}
        >
          {mensaje.texto}
        </div>
      )}

      {!puedeEditar && (
        <div className="card p-3 text-xs" style={{ color: "var(--texto-suave)" }}>
          Estás con acceso de solo lectura: puedes ver el borrador, pero no puedes generar, editar ni aprobar nada.
        </div>
      )}

      {(proyectos?.length ?? 0) === 0 && puedeEditar && (
        <div className="card p-6 text-center">
          <p className="mb-3" style={{ color: "var(--texto-suave)" }}>
            Todavía no hay ningún borrador de planificación para {anioBorrador}.
          </p>
          <button className="boton-primario" onClick={generarBorrador} disabled={generando}>
            {generando ? "Generando…" : `Generar borrador ${anioBorrador ?? ""}`}
          </button>
        </div>
      )}

      {(proyectos?.length ?? 0) > 0 && (
        <>
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
              {puedeEditar && (
                <div className="flex items-center gap-2 ml-auto">
                  <button className="boton-secundario" onClick={() => setMostrarFormNuevo((v) => !v)}>
                    {mostrarFormNuevo ? "Cancelar" : "+ Agregar proyecto"}
                  </button>
                  <button className="boton-secundario" onClick={generarBorrador} disabled={generando}>
                    {generando ? "Regenerando…" : "Regenerar borrador"}
                  </button>
                  <button className="boton-primario" onClick={cargar} disabled={cargando}>
                    {cargando ? "Actualizando…" : "Actualizar"}
                  </button>
                </div>
              )}
            </div>
            {actualizadoEn && (
              <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                Actualizado {new Date(actualizadoEn).toLocaleString("es-PE")}
              </span>
            )}
          </div>

          {mostrarFormNuevo && (
            <FormularioNuevoProyectoBorrador
              onCancelar={() => setMostrarFormNuevo(false)}
              onCreado={() => {
                setMostrarFormNuevo(false);
                cargar();
              }}
            />
          )}

          <div className="card p-0 overflow-hidden">
            <div style={{ maxHeight: "65vh", overflow: "auto" }}>
              <table className="text-xs border-collapse" style={{ tableLayout: "fixed", width: 1660 }}>
                <colgroup>
                  <col style={{ width: 40 }} />
                  <col style={{ width: 190 }} />
                  <col style={{ width: 190 }} />
                  <col style={{ width: 90 }} />
                  <col style={{ width: 115 }} />
                  {NOMBRES_MES.map((m) => (
                    <col key={m} style={{ width: 90 }} />
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
                    {NOMBRES_MES.map((m) => (
                      <th key={m} className="py-2 px-3 font-semibold text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                        {m}
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
                        <td className="py-1.5 px-3 text-center" style={ESTILO_DESTACADO}>
                          <CampoEditable
                            fila={p.filaExcel}
                            campo="presupuestoAprobado"
                            tipo="moneda"
                            valor={p.presupuestoAprobado}
                            endpoint={ENDPOINT_CELDA}
                            soloLectura={!puedeEditar}
                            className="text-center whitespace-nowrap font-medium"
                            onGuardado={(v) => actualizarLocal(p.filaExcel, { presupuestoAprobado: Number(v) })}
                          />
                          {mostrarSoles && (
                            <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                              {soles(p.presupuestoAprobado, tipoCambio)}
                            </span>
                          )}
                        </td>
                        {NOMBRES_MES.map((_, mi) => (
                          <td key={mi} className="py-1.5 px-3" style={{ borderLeft: "1px solid var(--borde)" }}>
                            <CampoMontoSumado
                              fila={p.filaExcel}
                              campo={`proyectado:${mi}`}
                              valor={p.proyectado[mi]}
                              className="text-center text-xs"
                              endpoint={ENDPOINT_CELDA}
                              soloLectura={!puedeEditar}
                              mostrarSoles={mostrarSoles}
                              tipoCambio={tipoCambio}
                              onGuardado={(nuevo) =>
                                actualizarLocal(p.filaExcel, { proyectado: p.proyectado.map((x, j) => (j === mi ? nuevo : x)) })
                              }
                            />
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
                              ? "Ya se repartió más de lo aprobado"
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
                      <td colSpan={18} className="py-6 text-center" style={{ color: "var(--texto-suave)" }}>
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
                    <td colSpan={12} style={{ background: "var(--bg)" }}></td>
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

          {puedeEditar && (
            <div className="card p-4" style={{ background: "#fbe1ec", border: "1px solid var(--peligro)" }}>
              <h2 className="font-semibold mb-1">Aprobar y activar {anioBorrador}</h2>
              <p className="text-sm mb-3" style={{ color: "var(--texto-suave)" }}>
                Cuando el presupuesto {anioBorrador} ya esté revisado y aprobado, este botón lo hace el nuevo
                BD_CAPEX en vivo — de ahí en adelante todos los indicadores de la app (Dashboard, Detalle,
                Facturas) van a salir de acá. El BD_CAPEX de {anioActivoApp} queda archivado con el año en el
                nombre, disponible para siempre, pero deja de ser el que se ve en pantalla.
              </p>
              <button className="boton-primario" onClick={aprobarYActivar} disabled={aprobando}>
                {aprobando ? "Activando…" : `Aprobar y activar ${anioBorrador}`}
              </button>
            </div>
          )}
        </>
      )}
    </div>
  );
}

function FormularioNuevoProyectoBorrador({ onCancelar, onCreado }: { onCancelar: () => void; onCreado: () => void }) {
  const [proyecto, setProyecto] = useState("");
  const [subNegocio, setSubNegocio] = useState("");
  const [grupoNegocio, setGrupoNegocio] = useState("");
  const [detalle, setDetalle] = useState("");
  const [prioridad, setPrioridad] = useState("");
  const [presupuestoAprobado, setPresupuestoAprobado] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const listo = grupoNegocio && prioridad.trim() && presupuestoAprobado.trim();

  async function crear() {
    setError(null);
    setEnviando(true);
    try {
      const res = await fetch("/api/capex/planificacion/agregar-proyecto", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          proyecto: proyecto.trim(),
          subNegocio: subNegocio.trim(),
          grupoNegocio,
          detalle: detalle.trim(),
          prioridad: prioridad.trim(),
          presupuestoAprobado: Number(presupuestoAprobado),
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo crear el proyecto.");
      onCreado();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setEnviando(false);
    }
  }

  return (
    <div className="card p-4" style={{ background: "var(--acento-suave)" }}>
      <h2 className="font-semibold mb-3 text-sm">Nuevo proyecto para el borrador</h2>
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div>
          <span className="etiqueta mb-0">Proyecto</span>
          <input className="campo" value={proyecto} onChange={(e) => setProyecto(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Sub Negocio</span>
          <input className="campo" value={subNegocio} onChange={(e) => setSubNegocio(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Grupo de Negocio *</span>
          <select className="campo" value={grupoNegocio} onChange={(e) => setGrupoNegocio(e.target.value)} disabled={enviando}>
            <option value="">Elige uno…</option>
            {GRUPOS_NEGOCIO.map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
          </select>
        </div>
        <div className="sm:col-span-2">
          <span className="etiqueta mb-0">Detalle</span>
          <input className="campo" value={detalle} onChange={(e) => setDetalle(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Prioridad *</span>
          <input className="campo" value={prioridad} onChange={(e) => setPrioridad(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Presupuesto Aprobado (USD) *</span>
          <input
            type="number"
            className="campo"
            value={presupuestoAprobado}
            onChange={(e) => setPresupuestoAprobado(e.target.value)}
            disabled={enviando}
          />
        </div>
      </div>
      {error && (
        <p className="text-sm mt-3" style={{ color: "var(--peligro)" }}>
          {error}
        </p>
      )}
      <div className="flex gap-2 mt-4">
        <button className="boton-primario" onClick={crear} disabled={!listo || enviando}>
          {enviando ? "Creando…" : "Crear proyecto"}
        </button>
        <button className="boton-secundario" onClick={onCancelar} disabled={enviando}>
          Cancelar
        </button>
      </div>
    </div>
  );
}
