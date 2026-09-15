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
const ENDPOINT_CELDA = "/api/opex/planificacion/celda";

function normalizar(texto: string): string {
  return texto.toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
}

/** Umbral para considerar una línea "programada": menos de medio centavo de diferencia
 *  entre el Presupuesto Aprobado y lo ya repartido entre los 12 meses. */
const TOLERANCIA_DIFERENCIA = 0.005;

const ESTILO_DESTACADO = { background: "var(--acento-suave)" };

/**
 * Planificación OPEX: mismo criterio que Planificación CAPEX — el borrador del
 * presupuesto del año SIGUIENTE al activo vive en su propia hoja ("Presupuesto
 * (Planificación)"), separada de la hoja en vivo, hasta que se aprueba y activa.
 */
export default function PlanificacionOpex() {
  const nivelAcceso = useNivelAcceso();
  const puedeEditar = nivelAcceso === "completo";
  const anioActivoApp = useAnios().opex;

  const [lineas, setLineas] = useState<ProyectoCapex[] | null>(null);
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
  const [grupoSel, setGrupoSel] = usePersistedState<string | null>("planificacion-opex-grupo", null);
  const [soloSinProgramar, setSoloSinProgramar] = usePersistedState("planificacion-opex-sin-programar", false);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("planificacion-opex-soles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();

  function actualizarLocal(fila: number, cambios: Partial<ProyectoCapex>) {
    setLineas((prev) => prev?.map((x) => (x.filaExcel === fila ? { ...x, ...cambios } : x)) ?? prev);
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/opex/planificacion", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setLineas(json.lineas);
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
      (lineas?.length ?? 0) > 0 &&
      !window.confirm(
        "Ya hay un borrador con datos. Generarlo de nuevo BORRA todo lo que se haya escrito acá (vuelve a copiar la lista de líneas vigente, en blanco). ¿Continuar?"
      )
    ) {
      return;
    }
    setGenerando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/opex/planificacion/generar", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo generar el borrador.");
      setMensaje({ tipo: "ok", texto: `Borrador generado con ${json.lineas} líneas.` });
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
        `¿Aprobar y activar ${anioBorrador}? Desde este momento el Presupuesto OPEX en vivo (y todos los indicadores de la app) van a ser los de este borrador. El presupuesto de ${anioActivoApp} queda archivado con el año en el nombre — no se pierde, pero deja de ser el que se ve en el Dashboard. Esta acción no se puede deshacer desde la app.`
      )
    ) {
      return;
    }
    setAprobando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/opex/planificacion/aprobar", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo aprobar y activar.");
      setMensaje({
        tipo: "ok",
        texto: `${json.anioNuevo - 1} archivado como "${json.hojaArchivada}". El presupuesto ${json.anioNuevo} ya está en vivo.`,
      });
      await cargar();
      window.location.href = "/opex";
    } catch (e) {
      setMensaje({ tipo: "error", texto: (e as Error).message });
    } finally {
      setAprobando(false);
    }
  }

  const resueltas = useMemo(() => resolverProyectos(lineas ?? [], 0), [lineas]);

  const grupos = useMemo(
    () => Array.from(new Set(resueltas.map((l) => l.grupoNegocio || "SIN GRUPO"))).sort((a, b) => a.localeCompare(b, "es")),
    [resueltas]
  );

  const sinProgramarTotal = useMemo(
    () => resueltas.filter((l) => Math.abs(l.diferencia) > TOLERANCIA_DIFERENCIA).length,
    [resueltas]
  );

  const filtradas = useMemo(() => {
    let filas = resueltas;
    if (grupoSel) filas = filas.filter((l) => (l.grupoNegocio || "SIN GRUPO") === grupoSel);
    if (soloSinProgramar) filas = filas.filter((l) => Math.abs(l.diferencia) > TOLERANCIA_DIFERENCIA);
    const q = normalizar(busqueda);
    if (q) filas = filas.filter((l) => normalizar(l.proyecto).includes(q) || normalizar(l.detalle).includes(q));
    return [...filas].sort(
      (a, b) => a.grupoNegocio.localeCompare(b.grupoNegocio, "es") || a.proyecto.localeCompare(b.proyecto, "es")
    );
  }, [resueltas, grupoSel, soloSinProgramar, busqueda]);

  const totales = useMemo(
    () =>
      filtradas.reduce(
        (acc, l) => ({ presupuestoAprobado: acc.presupuestoAprobado + l.presupuestoAprobado, diferencia: acc.diferencia + l.diferencia }),
        { presupuestoAprobado: 0, diferencia: 0 }
      ),
    [filtradas]
  );

  if (cargando && lineas === null) {
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
        <h2 className="text-lg font-semibold">Planificación OPEX {anioBorrador ?? ""}</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Borrador del presupuesto {anioBorrador} — separado del Presupuesto en vivo ({anioActivoApp},{" "}
          {archivo || "…"}), no afecta nada hasta que se apruebe. {filtradas.length} de {lineas?.length ?? 0}{" "}
          líneas.
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

      {(lineas?.length ?? 0) === 0 && puedeEditar && (
        <div className="card p-6 text-center">
          <p className="mb-3" style={{ color: "var(--texto-suave)" }}>
            Todavía no hay ningún borrador de planificación para {anioBorrador}.
          </p>
          <button className="boton-primario" onClick={generarBorrador} disabled={generando}>
            {generando ? "Generando…" : `Generar borrador ${anioBorrador ?? ""}`}
          </button>
        </div>
      )}

      {(lineas?.length ?? 0) > 0 && (
        <>
          <div className="card p-4 flex flex-col gap-3">
            <div className="flex flex-wrap items-center gap-3">
              <input
                type="text"
                placeholder="Buscar línea o detalle…"
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
                title="Muestra solo las líneas donde el Presupuesto Aprobado todavía no está totalmente repartido en los 12 meses"
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
                    {mostrarFormNuevo ? "Cancelar" : "+ Agregar línea"}
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
            <FormularioNuevaLineaBorrador
              onCancelar={() => setMostrarFormNuevo(false)}
              onCreada={() => {
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
                    <th className="py-2 px-3 font-semibold">Línea de gasto</th>
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
                  {filtradas.map((l, i) => {
                    const sinProgramar = Math.abs(l.diferencia) > TOLERANCIA_DIFERENCIA;
                    return (
                      <tr
                        key={l.filaExcel}
                        style={{ borderTop: "1px solid var(--borde)", background: sinProgramar ? "#fff8e6" : undefined }}
                      >
                        <td className="py-1.5 px-3 text-center" style={{ color: "var(--texto-suave)" }}>
                          {i + 1}
                        </td>
                        <td className="py-1.5 px-3 truncate" title={l.proyecto}>
                          {l.proyecto || "—"}
                        </td>
                        <td className="py-1.5 px-3">
                          <CampoEditable
                            fila={l.filaExcel}
                            campo="detalle"
                            tipo="texto"
                            valor={l.detalle}
                            placeholder="—"
                            endpoint={ENDPOINT_CELDA}
                            soloLectura={!puedeEditar}
                            onGuardado={(v) => actualizarLocal(l.filaExcel, { detalle: String(v) })}
                          />
                        </td>
                        <td className="py-1.5 px-3 truncate">{l.grupoNegocio}</td>
                        <td className="py-1.5 px-3 text-center" style={ESTILO_DESTACADO}>
                          <CampoEditable
                            fila={l.filaExcel}
                            campo="presupuestoAprobado"
                            tipo="moneda"
                            valor={l.presupuestoAprobado}
                            endpoint={ENDPOINT_CELDA}
                            soloLectura={!puedeEditar}
                            className="text-center whitespace-nowrap font-medium"
                            onGuardado={(v) => actualizarLocal(l.filaExcel, { presupuestoAprobado: Number(v) })}
                          />
                          {mostrarSoles && (
                            <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                              {soles(l.presupuestoAprobado, tipoCambio)}
                            </span>
                          )}
                        </td>
                        {NOMBRES_MES.map((_, mi) => (
                          <td key={mi} className="py-1.5 px-3" style={{ borderLeft: "1px solid var(--borde)" }}>
                            <CampoMontoSumado
                              fila={l.filaExcel}
                              campo={`proyectado:${mi}`}
                              valor={l.proyectado[mi]}
                              className="text-center text-xs"
                              endpoint={ENDPOINT_CELDA}
                              soloLectura={!puedeEditar}
                              mostrarSoles={mostrarSoles}
                              tipoCambio={tipoCambio}
                              onGuardado={(nuevo) =>
                                actualizarLocal(l.filaExcel, { proyectado: l.proyectado.map((x, j) => (j === mi ? nuevo : x)) })
                              }
                            />
                          </td>
                        ))}
                        <td
                          className="py-1.5 px-3 text-center whitespace-nowrap font-bold"
                          style={{
                            borderLeft: "1px solid var(--borde)",
                            color: l.diferencia < -TOLERANCIA_DIFERENCIA
                              ? "var(--peligro)"
                              : l.diferencia > TOLERANCIA_DIFERENCIA
                                ? "var(--alerta)"
                                : "var(--exito)",
                          }}
                          title={
                            l.diferencia < -TOLERANCIA_DIFERENCIA
                              ? "Ya se repartió más de lo aprobado"
                              : l.diferencia > TOLERANCIA_DIFERENCIA
                                ? "Todavía queda presupuesto sin repartir en algún mes"
                                : "Presupuesto totalmente programado"
                          }
                        >
                          {moneda2(l.diferencia)}
                        </td>
                      </tr>
                    );
                  })}
                  {filtradas.length === 0 && (
                    <tr>
                      <td colSpan={18} className="py-6 text-center" style={{ color: "var(--texto-suave)" }}>
                        Sin líneas para este filtro.
                      </td>
                    </tr>
                  )}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: "2px solid var(--borde)" }}>
                    <td className="py-2 px-3 font-bold" colSpan={4} style={{ background: "var(--bg)" }}>
                      Total ({filtradas.length})
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
                Presupuesto OPEX en vivo — de ahí en adelante todos los indicadores de la app van a salir de acá.
                El presupuesto de {anioActivoApp} queda archivado con el año en el nombre, disponible para
                siempre, pero deja de ser el que se ve en pantalla.
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

function FormularioNuevaLineaBorrador({ onCancelar, onCreada }: { onCancelar: () => void; onCreada: () => void }) {
  const [empresa, setEmpresa] = useState("");
  const [grupoGasto, setGrupoGasto] = useState("");
  const [subgrupoGasto, setSubgrupoGasto] = useState("");
  const [lineaGasto, setLineaGasto] = useState("");
  const [moneda, setMoneda] = useState("USD");
  const [detalle, setDetalle] = useState("");
  const [responsable, setResponsable] = useState("");
  const [presupuestoAprobado, setPresupuestoAprobado] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const listo = empresa.trim() && grupoGasto.trim() && subgrupoGasto.trim() && lineaGasto.trim() && presupuestoAprobado.trim();

  async function crear() {
    setError(null);
    setEnviando(true);
    try {
      const res = await fetch("/api/opex/planificacion/agregar-linea", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          empresa: empresa.trim(),
          grupoGasto: grupoGasto.trim(),
          subgrupoGasto: subgrupoGasto.trim(),
          lineaGasto: lineaGasto.trim(),
          moneda,
          detalle: detalle.trim(),
          responsable: responsable.trim(),
          presupuestoAprobado: Number(presupuestoAprobado),
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo crear la línea.");
      onCreada();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setEnviando(false);
    }
  }

  return (
    <div className="card p-4" style={{ background: "var(--acento-suave)" }}>
      <h2 className="font-semibold mb-3 text-sm">Nueva línea para el borrador</h2>
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div>
          <span className="etiqueta mb-0">Empresa *</span>
          <input className="campo" value={empresa} onChange={(e) => setEmpresa(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Grupo de Gasto *</span>
          <input className="campo" value={grupoGasto} onChange={(e) => setGrupoGasto(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Subgrupo de Gasto *</span>
          <input className="campo" value={subgrupoGasto} onChange={(e) => setSubgrupoGasto(e.target.value)} disabled={enviando} />
        </div>
        <div className="sm:col-span-2">
          <span className="etiqueta mb-0">Línea de Gasto *</span>
          <input className="campo" value={lineaGasto} onChange={(e) => setLineaGasto(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Moneda</span>
          <select className="campo" value={moneda} onChange={(e) => setMoneda(e.target.value)} disabled={enviando}>
            <option value="USD">USD</option>
            <option value="Soles">Soles</option>
          </select>
        </div>
        <div className="sm:col-span-2">
          <span className="etiqueta mb-0">Detalle</span>
          <input className="campo" value={detalle} onChange={(e) => setDetalle(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Responsable</span>
          <input className="campo" value={responsable} onChange={(e) => setResponsable(e.target.value)} disabled={enviando} />
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
          {enviando ? "Creando…" : "Crear línea"}
        </button>
        <button className="boton-secundario" onClick={onCancelar} disabled={enviando}>
          Cancelar
        </button>
      </div>
    </div>
  );
}
