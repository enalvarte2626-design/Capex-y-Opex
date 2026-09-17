"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";
import { moneda2 } from "@/lib/format";

const ROSA_TEMPLATE = "#d6246e";

interface ArchivoParaComparar {
  nombre: string;
  itemId: string;
  esArchivoActual: boolean;
  cerradosSugeridos: number;
}

interface VersionArchivo {
  id: string;
  fecha: string;
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

interface ResumenGrupo {
  grupoNegocio: string;
  presupuestoAprobadoAntes: number;
  presupuestoAprobadoAhora: number;
  forecastAntes: number;
  forecastAhora: number;
  diferenciaAntes: number;
  diferenciaAhora: number;
  cambio: number;
}

interface ResultadoComparacion {
  nombreArchivoAnterior: string;
  fechaVersionAnterior: string | null;
  cerradosEnReferencia: number;
  cerradosActual: number;
  proyectosNuevos: string[];
  cambios: CambioLinea[];
  resumenPorGrupo: ResumenGrupo[];
  totalAntes: number;
  totalAhora: number;
  totalCambio: number;
}

const VERSION_ACTUAL = "__actual__";

/**
 * Compara BD_CAPEX en vivo contra un punto de referencia elegido a mano: otro archivo de
 * la misma carpeta, o una VERSIÓN ANTERIOR de SharePoint de cualquiera de ellos (incluido
 * el mismo archivo en vivo). Esto último es clave cuando el archivo que se "cerró" y
 * presentó siguió editándose después (ej. se le agregaron montos de un mes nuevo antes de
 * generar el archivo siguiente): la comparación no debe hacerse contra el contenido MÁS
 * RECIENTE de ese archivo (ya trae esos cambios), sino contra la versión de SharePoint
 * guardada en el momento real de cierre/presentación.
 */
export default function CompararCierre() {
  const [archivos, setArchivos] = useState<ArchivoParaComparar[] | null>(null);
  const [archivoActualNombre, setArchivoActualNombre] = useState("");
  const [cargandoLista, setCargandoLista] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [itemIdSel, setItemIdSel] = useState("");
  const [versiones, setVersiones] = useState<VersionArchivo[] | null>(null);
  const [cargandoVersiones, setCargandoVersiones] = useState(false);
  const [versionSel, setVersionSel] = useState(VERSION_ACTUAL);
  const [cerrados, setCerrados] = useState(0);
  const [cerradosActual, setCerradosActual] = useState(0);

  const [comparando, setComparando] = useState(false);
  const [resultado, setResultado] = useState<ResultadoComparacion | null>(null);
  const [soloMesesCerrados, setSoloMesesCerrados] = useState(false);
  const [grupoSel, setGrupoSel] = useState("");
  const [soloPrioridad12, setSoloPrioridad12] = useState(false);

  useEffect(() => {
    fetch("/api/capex/cierres-disponibles", { cache: "no-store" })
      .then((r) => r.json())
      .then((j) => {
        if (j.error) throw new Error(j.error);
        setArchivos(j.disponibles);
        setArchivoActualNombre(j.archivoActual);
        // Por defecto, el primer archivo que NO es el actual (si hay alguno) — lo más
        // típico es comparar el archivo en vivo contra otro archivo de un cierre pasado.
        const sugerido = (j.disponibles as ArchivoParaComparar[]).find((a) => !a.esArchivoActual);
        if (sugerido) {
          setItemIdSel(sugerido.itemId);
          setCerrados(sugerido.cerradosSugeridos);
        }
        const actual = (j.disponibles as ArchivoParaComparar[]).find((a) => a.esArchivoActual);
        if (actual) setCerradosActual(actual.cerradosSugeridos);
      })
      .catch((e) => setError((e as Error).message))
      .finally(() => setCargandoLista(false));
  }, []);

  const archivoSel = useMemo(() => archivos?.find((a) => a.itemId === itemIdSel) ?? null, [archivos, itemIdSel]);

  useEffect(() => {
    if (!itemIdSel) return;
    setVersiones(null);
    setVersionSel(VERSION_ACTUAL);
    setCargandoVersiones(true);
    fetch(`/api/capex/versiones-archivo?itemId=${encodeURIComponent(itemIdSel)}`, { cache: "no-store" })
      .then((r) => r.json())
      .then((j) => {
        if (j.error) throw new Error(j.error);
        setVersiones(j.versiones);
      })
      .catch((e) => setError((e as Error).message))
      .finally(() => setCargandoVersiones(false));
  }, [itemIdSel]);

  async function comparar() {
    if (!archivoSel) return;
    setComparando(true);
    setError(null);
    setResultado(null);
    try {
      const version = versionSel !== VERSION_ACTUAL ? versiones?.find((v) => v.id === versionSel) : null;
      const params = new URLSearchParams({
        itemId: archivoSel.itemId,
        nombre: archivoSel.nombre,
        cerrados: String(cerrados),
        cerradosActual: String(cerradosActual),
      });
      if (soloPrioridad12) params.set("prioridades", "1,2");
      if (version) {
        params.set("versionId", version.id);
        params.set("fechaVersion", version.fecha);
      }
      const res = await fetch(`/api/capex/comparar-cierre?${params.toString()}`, { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo comparar.");
      setResultado(json);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setComparando(false);
    }
  }

  const gruposDisponibles = useMemo(
    () => Array.from(new Set((resultado?.cambios ?? []).map((c) => c.grupoNegocio || "SIN GRUPO"))).sort((a, b) => a.localeCompare(b, "es")),
    [resultado]
  );

  const cambiosMostrados = useMemo(() => {
    if (!resultado) return [];
    let filas = resultado.cambios;
    if (soloMesesCerrados) filas = filas.filter((c) => c.eraMesCerrado);
    if (grupoSel) filas = filas.filter((c) => (c.grupoNegocio || "SIN GRUPO") === grupoSel);
    return filas;
  }, [resultado, soloMesesCerrados, grupoSel]);

  const totalMesesCerrados = resultado?.cambios.filter((c) => c.eraMesCerrado).length ?? 0;

  return (
    <div className="flex flex-col gap-4">
      <div>
        <h2 className="text-lg font-semibold">Comparar con un cierre anterior</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Compara BD_CAPEX en vivo ({archivoActualNombre || "…"}) contra otro archivo de la carpeta, o contra una
          versión anterior de SharePoint de cualquiera de ellos — útil cuando el archivo que se presentó siguió
          editándose después, y su contenido más reciente ya no es el que se cerró ese día.
        </p>
      </div>

      {cargandoLista && <p style={{ color: "var(--texto-suave)" }}>Buscando archivos en la misma carpeta…</p>}

      {error && (
        <div className="card p-3 text-sm" style={{ background: "#fbe1ec", color: "var(--peligro)" }}>
          {error}
        </div>
      )}

      {!cargandoLista && archivos && archivos.length > 0 && (
        <div className="card p-4 flex flex-col gap-3">
          <div className="flex flex-wrap items-center gap-3">
            <span className="etiqueta mb-0">Archivo:</span>
            <select className="campo" style={{ width: "auto", minWidth: 280 }} value={itemIdSel} onChange={(e) => {
              const a = archivos.find((x) => x.itemId === e.target.value);
              setItemIdSel(e.target.value);
              setCerrados(a?.cerradosSugeridos ?? 0);
            }}>
              {archivos.map((a) => (
                <option key={a.itemId} value={a.itemId}>
                  {a.nombre}
                  {a.esArchivoActual ? " (en vivo)" : ""}
                </option>
              ))}
            </select>

            <span className="etiqueta mb-0">Versión:</span>
            <select
              className="campo"
              style={{ width: "auto", minWidth: 260 }}
              value={versionSel}
              onChange={(e) => setVersionSel(e.target.value)}
              disabled={cargandoVersiones || !versiones}
            >
              <option value={VERSION_ACTUAL}>Más reciente (contenido actual del archivo)</option>
              {versiones?.map((v, i) => (
                <option key={v.id} value={v.id}>
                  {new Date(v.fecha).toLocaleString("es-PE")}
                  {i === versiones.length - 1 ? " (primera versión guardada)" : ""}
                </option>
              ))}
            </select>
            {cargandoVersiones && (
              <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                Buscando versiones…
              </span>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <span className="etiqueta mb-0">Meses ya cerrados en ese punto:</span>
            <select className="campo" style={{ width: "auto" }} value={cerrados} onChange={(e) => setCerrados(Number(e.target.value))}>
              <option value={0}>Ninguno</option>
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  Hasta {nombre}
                </option>
              ))}
            </select>
            <span className="text-xs" style={{ color: "var(--texto-suave)" }} title="No siempre coincide con lo que dice el nombre del archivo — ajústalo si ese archivo siguió editándose después de generarse.">
              (ajústalo si el archivo se siguió editando después de &quot;cerrarse&quot;)
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <span className="etiqueta mb-0">Meses cerrados ahora:</span>
            <select
              className="campo"
              style={{ width: "auto" }}
              value={cerradosActual}
              onChange={(e) => setCerradosActual(Number(e.target.value))}
            >
              <option value={0}>Ninguno</option>
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  Hasta {nombre}
                </option>
              ))}
            </select>
            <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
              — para calcular el resumen por Grupo de Negocio (abajo)
            </span>
            <span
              className="chip"
              data-activo={soloPrioridad12}
              onClick={() => setSoloPrioridad12((v) => !v)}
              title="Deja completamente afuera de la comparación (cambios, proyectos nuevos y resumen por grupo) cualquier proyecto que no sea Prioridad 1 o 2"
            >
              {soloPrioridad12 ? "✓ " : ""}Solo Prioridad 1 y 2
            </span>
            <button className="boton-primario ml-auto" onClick={comparar} disabled={comparando || !archivoSel}>
              {comparando ? "Comparando…" : "Comparar"}
            </button>
          </div>
        </div>
      )}

      {resultado && (
        <>
          <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
            Comparado contra: <strong>{resultado.nombreArchivoAnterior}</strong>
            {resultado.fechaVersionAnterior && ` — versión del ${new Date(resultado.fechaVersionAnterior).toLocaleString("es-PE")}`}
            {" · "}
            {resultado.cerradosEnReferencia} mes(es) cerrado(s) en ese punto, {resultado.cerradosActual} mes(es) cerrado(s) ahora.
          </p>

          <div className="card p-4">
            <h3 className="font-semibold mb-3">Diferencia presupuestal: antes vs. ahora</h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
              <div className="rounded-lg p-3" style={{ background: "var(--bg)" }}>
                <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
                  Diferencia en {resultado.nombreArchivoAnterior}
                </p>
                <p className="text-2xl font-bold" style={{ color: resultado.totalAntes < 0 ? "var(--peligro)" : "var(--exito)" }}>
                  {moneda2(resultado.totalAntes)}
                </p>
              </div>
              <div className="rounded-lg p-3" style={{ background: "var(--bg)" }}>
                <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
                  Diferencia ahora
                </p>
                <p className="text-2xl font-bold" style={{ color: resultado.totalAhora < 0 ? "var(--peligro)" : "var(--exito)" }}>
                  {moneda2(resultado.totalAhora)}
                </p>
              </div>
              <div className="rounded-lg p-3" style={{ background: resultado.totalCambio >= 0 ? "#e3f3e3" : "#fbe1ec" }}>
                <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
                  Cambio ({resultado.totalCambio >= 0 ? "mejoró / ahorro" : "empeoró / sobrepaso"})
                </p>
                <p className="text-2xl font-bold" style={{ color: resultado.totalCambio >= 0 ? "var(--exito)" : "var(--peligro)" }}>
                  {resultado.totalCambio > 0 ? "+" : ""}
                  {moneda2(resultado.totalCambio)}
                </p>
              </div>
            </div>

            <h4 className="font-semibold text-sm mb-2">Por Grupo de Negocio, ordenado por el cambio más grande</h4>
            <div className="overflow-x-auto">
              <table className="text-xs border-collapse w-full">
                <thead>
                  <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                    <th className="py-1.5 px-3 font-semibold" rowSpan={2}>
                      Grupo
                    </th>
                    <th className="py-1.5 px-3 font-semibold text-center" colSpan={2} style={{ borderLeft: "1px solid var(--borde)" }}>
                      Presupuesto Aprobado
                    </th>
                    <th className="py-1.5 px-3 font-semibold text-center" colSpan={2} style={{ borderLeft: "1px solid var(--borde)" }}>
                      Forecast
                    </th>
                    <th className="py-1.5 px-3 font-semibold text-center" colSpan={3} style={{ borderLeft: "1px solid var(--borde)" }}>
                      Diferencia
                    </th>
                  </tr>
                  <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                    <th className="py-1 px-3 font-normal text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                      Antes
                    </th>
                    <th className="py-1 px-3 font-normal text-center">Ahora</th>
                    <th className="py-1 px-3 font-normal text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                      Antes
                    </th>
                    <th className="py-1 px-3 font-normal text-center">Ahora</th>
                    <th className="py-1 px-3 font-normal text-center" style={{ borderLeft: "1px solid var(--borde)" }}>
                      Antes
                    </th>
                    <th className="py-1 px-3 font-normal text-center">Ahora</th>
                    <th className="py-1 px-3 font-normal text-center">Cambio</th>
                  </tr>
                </thead>
                <tbody>
                  {resultado.resumenPorGrupo.map((g) => (
                    <tr key={g.grupoNegocio} style={{ borderTop: "1px solid var(--borde)" }}>
                      <td className="py-1.5 px-3 font-semibold">{g.grupoNegocio}</td>
                      <td className="py-1.5 px-3 text-center whitespace-nowrap" style={{ borderLeft: "1px solid var(--borde)" }}>
                        {moneda2(g.presupuestoAprobadoAntes)}
                      </td>
                      <td className="py-1.5 px-3 text-center whitespace-nowrap">{moneda2(g.presupuestoAprobadoAhora)}</td>
                      <td className="py-1.5 px-3 text-center whitespace-nowrap" style={{ borderLeft: "1px solid var(--borde)" }}>
                        {moneda2(g.forecastAntes)}
                      </td>
                      <td className="py-1.5 px-3 text-center whitespace-nowrap">{moneda2(g.forecastAhora)}</td>
                      <td
                        className="py-1.5 px-3 text-center whitespace-nowrap"
                        style={{ borderLeft: "1px solid var(--borde)", color: g.diferenciaAntes < 0 ? "var(--peligro)" : undefined }}
                      >
                        {moneda2(g.diferenciaAntes)}
                      </td>
                      <td className="py-1.5 px-3 text-center whitespace-nowrap" style={{ color: g.diferenciaAhora < 0 ? "var(--peligro)" : undefined }}>
                        {moneda2(g.diferenciaAhora)}
                      </td>
                      <td
                        className="py-1.5 px-3 text-center whitespace-nowrap font-bold"
                        style={{ color: g.cambio >= 0 ? "var(--exito)" : "var(--peligro)" }}
                      >
                        {g.cambio > 0 ? "+" : ""}
                        {moneda2(g.cambio)} {g.cambio >= 0 ? "▲ ahorro" : "▼ gasto aumentado"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {resultado.proyectosNuevos.length > 0 && (
            <div className="card p-4 text-sm">
              <p className="font-semibold mb-1">{resultado.proyectosNuevos.length} proyecto(s) nuevo(s) desde entonces</p>
              <p style={{ color: "var(--texto-suave)" }}>
                No existían todavía en ese punto de referencia, así que no hay con qué compararlos:{" "}
                {resultado.proyectosNuevos.join(", ")}.
              </p>
            </div>
          )}

          {resultado.cambios.length === 0 ? (
            <div className="card p-6 text-center" style={{ color: "var(--exito)" }}>
              Ningún proyecto cambió (Presupuesto Aprobado, Real ni Proyectado) desde ese punto de referencia.
            </div>
          ) : (
            <div className="card p-0 overflow-hidden">
              <div className="p-4 flex items-center gap-3" style={{ borderBottom: "1px solid var(--borde)" }}>
                <span
                  className="chip"
                  data-activo={soloMesesCerrados}
                  onClick={() => setSoloMesesCerrados((v) => !v)}
                  title="Muestra solo cambios en un mes que ya estaba cerrado en la referencia — la causa más probable de un indicador que 'se mueve solo'"
                >
                  {soloMesesCerrados ? "✓ " : ""}Solo meses ya cerrados{totalMesesCerrados > 0 ? ` (${totalMesesCerrados})` : ""}
                </span>
                <select className="campo" style={{ width: "auto" }} value={grupoSel} onChange={(e) => setGrupoSel(e.target.value)}>
                  <option value="">Todos los grupos</option>
                  {gruposDisponibles.map((g) => (
                    <option key={g} value={g}>
                      {g}
                    </option>
                  ))}
                </select>
                <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                  {cambiosMostrados.length} de {resultado.cambios.length} cambios
                </span>
              </div>
              <div style={{ maxHeight: "65vh", overflow: "auto" }}>
                <table className="text-xs border-collapse w-full">
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)" }}>
                    <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                      <th className="py-2 px-3 font-semibold">Proyecto</th>
                      <th className="py-2 px-3 font-semibold">Grupo</th>
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
                        <td className="py-1.5 px-3">{c.grupoNegocio}</td>
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
