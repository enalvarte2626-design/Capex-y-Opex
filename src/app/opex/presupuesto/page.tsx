"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE, prioridadesDisponibles, resolverProyectos, type ProyectoCapex } from "@/lib/capex";
import { moneda2 } from "@/lib/format";
import { useMesCierre } from "@/lib/useMesCierre";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import CampoEditable from "@/components/CampoEditable";
import CampoMontoSumado from "@/components/CampoMontoSumado";
import ControlTipoCambio from "@/components/ControlTipoCambio";
import FiltroMultiple from "@/components/FiltroMultiple";

const NOMBRES_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const RANGO_DIACRITICOS = /[\u0300-\u036f]/g;

function normalizar(texto: string): string {
  return texto.toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
}

function soles(valorUsd: number, tipoCambio: number): string {
  return `S/ ${(valorUsd * tipoCambio).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

const ANCHO_SUBCOL_MES = 90;
const ESTILO_DESTACADO = { background: "var(--acento-suave)" };
/** N° + Empresa + Grupo + Subgrupo + Línea de Gasto + Responsable + Status, antes de los 4 importes destacados. */
const CANTIDAD_COLUMNAS_TEXTO = 7;

// Anchos fijos en px de las columnas que no son meses — sin esto (y sin
// table-layout:fixed), el navegador comprime las columnas de meses hasta dejar los
// importes cortados, porque el ancho puesto en cada <td> es solo una sugerencia en modo
// de layout automático. Con estos anchos + colgroup, quedan garantizados.
const COLUMNAS_BASE = [
  { clave: "numero", ancho: 44 },
  { clave: "empresa", ancho: 90 },
  { clave: "grupo", ancho: 160 },
  { clave: "subgrupo", ancho: 160 },
  { clave: "linea", ancho: 210 },
  { clave: "responsable", ancho: 140 },
  { clave: "status", ancho: 110 },
  { clave: "presupuesto", ancho: 150 },
  { clave: "gastoReal", ancho: 130 },
  { clave: "forecast", ancho: 130 },
  { clave: "diferencia", ancho: 130 },
] as const;
const ANCHO_TOTAL_BASE = COLUMNAS_BASE.reduce((a, c) => a + c.ancho, 0);

// Columnas "identificadoras" que quedan fijas al desplazarse hacia la derecha (mismo
// truco que "Inmovilizar paneles" de Excel) — hasta Línea de Gasto, para nunca perder de
// vista qué línea es mientras se scrollea hacia Responsable/Status/importes/meses.
const CLAVES_FIJAS = ["numero", "empresa", "grupo", "subgrupo", "linea"] as const;
const OFFSET_IZQUIERDA: Record<string, number> = {};
{
  let acumulado = 0;
  for (const c of COLUMNAS_BASE) {
    if (CLAVES_FIJAS.includes(c.clave as (typeof CLAVES_FIJAS)[number])) OFFSET_IZQUIERDA[c.clave] = acumulado;
    acumulado += c.ancho;
  }
}
function estiloFijo(clave: (typeof CLAVES_FIJAS)[number], fondo: string): React.CSSProperties {
  return {
    position: "sticky",
    left: OFFSET_IZQUIERDA[clave],
    zIndex: 6,
    background: fondo,
    boxShadow: clave === "linea" ? "2px 0 4px -2px rgba(0,0,0,0.15)" : undefined,
  };
}

export default function PresupuestoOpex() {
  const [lineas, setLineas] = useState<ProyectoCapex[] | null>(null);
  const [archivo, setArchivo] = useState("");
  const [actualizadoEn, setActualizadoEn] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [busqueda, setBusqueda] = useState("");
  const [gruposSel, setGruposSel] = usePersistedState<string[] | null>("opex-presupuesto-gruposSel", null);
  const [subgruposSel, setSubgruposSel] = usePersistedState<string[] | null>("opex-presupuesto-subgruposSel", null);
  const [statusSel, setStatusSel] = usePersistedState<string[] | null>("opex-presupuesto-statusSel", null);
  const [responsablesSel, setResponsablesSel] = usePersistedState<string[] | null>("opex-presupuesto-responsablesSel", null);
  const [mesesVisibles, setMesesVisibles] = usePersistedState("opex-presupuesto-mesesVisibles", false);
  const [mostrarFormNuevo, setMostrarFormNuevo] = useState(false);
  const [filaResaltada, setFilaResaltada] = useState<number | null>(null);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("opex-presupuesto-mostrarSoles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();
  const [mesCierre, setMesCierre] = useMesCierre();

  function actualizarLocal(fila: number, cambios: Partial<ProyectoCapex>) {
    setLineas((prev) => prev?.map((x) => (x.filaExcel === fila ? { ...x, ...cambios } : x)) ?? prev);
  }

  /** Guarda Status solo cuando el usuario elige un valor en el desplegable — así no se
   *  sobrescribe nada en filas que nadie tocó (varias ya tienen texto propio ahí, ej.
   *  "continuidad del 2025"). */
  async function guardarStatus(fila: number, valor: string) {
    const anterior = lineas?.find((x) => x.filaExcel === fila)?.status ?? "";
    actualizarLocal(fila, { status: valor });
    try {
      const res = await fetch("/api/opex/celda", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fila, campo: "status", valor }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
    } catch {
      actualizarLocal(fila, { status: anterior });
    }
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/opex", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setLineas(json.lineas);
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

  const resueltas = useMemo(() => resolverProyectos(lineas ?? [], mesCierre), [lineas, mesCierre]);

  // Siempre Real y Proyectado juntos, en todos los meses — incluidos los ya cerrados,
  // para poder comparar contra lo proyectado aunque el mes ya haya pasado. Solo el Real
  // de un mes cerrado queda de solo lectura (ver más abajo); el Proyectado se puede
  // seguir ajustando en cualquier mes.
  const mesesConSubcolumnas = useMemo(
    () =>
      NOMBRES_MES.map((nombre, mi) => ({
        nombre,
        mi,
        subcolumnas: ["real", "proyectado"] as const,
      })),
    []
  );
  const totalSubcolumnasMeses = mesesConSubcolumnas.reduce((a, m) => a + m.subcolumnas.length, 0);

  const grupos = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const l of resueltas) conteo.set(l.grupoNegocio, (conteo.get(l.grupoNegocio) ?? 0) + 1);
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([valor, cantidad]) => ({ valor, cantidad }));
  }, [resueltas]);
  const subgrupos = useMemo(() => prioridadesDisponibles(resueltas), [resueltas]);
  const estatus = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const l of resueltas) {
      const valor = l.status.trim() || "Sin definir";
      conteo.set(valor, (conteo.get(valor) ?? 0) + 1);
    }
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([valor, cantidad]) => ({ valor, cantidad }));
  }, [resueltas]);
  const responsables = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const l of resueltas) {
      const valor = l.responsable.trim() || "Sin responsable";
      conteo.set(valor, (conteo.get(valor) ?? 0) + 1);
    }
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([valor, cantidad]) => ({ valor, cantidad }));
  }, [resueltas]);

  // Al cargar por primera vez, arranca con todo seleccionado.
  useEffect(() => {
    if (gruposSel === null && grupos.length > 0) setGruposSel(grupos.map((g) => g.valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [grupos]);
  useEffect(() => {
    if (subgruposSel === null && subgrupos.length > 0) setSubgruposSel(subgrupos.map((s) => s.valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [subgrupos]);
  useEffect(() => {
    if (statusSel === null && estatus.length > 0) setStatusSel(estatus.map((s) => s.valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [estatus]);
  useEffect(() => {
    if (responsablesSel === null && responsables.length > 0) setResponsablesSel(responsables.map((r) => r.valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [responsables]);

  const filtradas = useMemo(() => {
    let filas = resueltas;
    if (gruposSel) filas = filas.filter((l) => gruposSel.includes(l.grupoNegocio));
    if (subgruposSel) filas = filas.filter((l) => subgruposSel.includes(l.prioridad || "Sin prioridad"));
    if (statusSel) filas = filas.filter((l) => statusSel.includes(l.status.trim() || "Sin definir"));
    if (responsablesSel) filas = filas.filter((l) => responsablesSel.includes(l.responsable.trim() || "Sin responsable"));
    const q = normalizar(busqueda);
    if (q) {
      filas = filas.filter(
        (l) => normalizar(l.proyecto).includes(q) || normalizar(l.detalle).includes(q) || normalizar(l.responsable).includes(q)
      );
    }
    return [...filas].sort(
      (a, b) => a.grupoNegocio.localeCompare(b.grupoNegocio, "es") || a.proyecto.localeCompare(b.proyecto, "es")
    );
  }, [resueltas, gruposSel, subgruposSel, statusSel, responsablesSel, busqueda]);

  const totales = useMemo(
    () =>
      filtradas.reduce(
        (acc, l) => ({
          presupuestoAprobado: acc.presupuestoAprobado + l.presupuestoAprobado,
          gastoReal: acc.gastoReal + l.gastoReal,
          forecast: acc.forecast + l.forecast,
          diferencia: acc.diferencia + l.diferencia,
        }),
        { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 }
      ),
    [filtradas]
  );

  if (cargando && !lineas) {
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
        <h2 className="text-lg font-semibold">Presupuesto OPEX 2026</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Todas las líneas de gasto de &quot;{archivo || "Presupuesto 2026"}&quot;, en vivo. {filtradas.length} de{" "}
          {resueltas.length} líneas.
        </p>
      </div>

      <div className="flex flex-wrap items-center gap-3">
        <input
          type="text"
          placeholder="Buscar línea de gasto…"
          className="campo"
          style={{ maxWidth: 280 }}
          value={busqueda}
          onChange={(e) => setBusqueda(e.target.value)}
        />
        <FiltroMultiple etiqueta="Grupo de Gasto" opciones={grupos} seleccion={gruposSel ?? []} onCambiar={setGruposSel} />
        <FiltroMultiple
          etiqueta="Subgrupo"
          opciones={subgrupos.map((s) => ({ valor: s.valor, cantidad: s.cantidad }))}
          seleccion={subgruposSel ?? []}
          onCambiar={setSubgruposSel}
        />
        <FiltroMultiple etiqueta="Status" opciones={estatus} seleccion={statusSel ?? []} onCambiar={setStatusSel} />
        <FiltroMultiple
          etiqueta="Responsable"
          opciones={responsables}
          seleccion={responsablesSel ?? []}
          onCambiar={setResponsablesSel}
        />
        {(busqueda || gruposSel || subgruposSel || statusSel || responsablesSel) && (
          <button
            type="button"
            className="text-xs font-semibold hover:underline"
            style={{ color: "var(--acento)" }}
            onClick={() => {
              setBusqueda("");
              setGruposSel(null);
              setSubgruposSel(null);
              setStatusSel(null);
              setResponsablesSel(null);
            }}
            title="Vuelve a mostrar todas las líneas, sin ningún filtro de por medio"
          >
            Quitar filtros ✕
          </button>
        )}
        <span
          className="chip"
          data-activo={mesesVisibles}
          onClick={() => setMesesVisibles((v) => !v)}
          title="Muestra u oculta el detalle mensual (Ene-Dic) de cada línea"
        >
          {mesesVisibles ? "Ocultar meses ▲" : "Ver meses ▾"}
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
          <select className="campo" style={{ width: "auto" }} value={mesCierre} onChange={(e) => setMesCierre(Number(e.target.value))}>
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
          <button className="boton-secundario" onClick={() => setMostrarFormNuevo((v) => !v)}>
            {mostrarFormNuevo ? "Cancelar" : "+ Agregar línea"}
          </button>
          <button className="boton-primario" onClick={cargar} disabled={cargando}>
            {cargando ? "Actualizando…" : "Actualizar"}
          </button>
        </div>
      </div>

      {mostrarFormNuevo && (
        <FormularioNuevaLinea
          gruposExistentes={grupos.map((g) => g.valor)}
          onCancelar={() => setMostrarFormNuevo(false)}
          onCreada={() => {
            setMostrarFormNuevo(false);
            cargar();
          }}
        />
      )}

      <div className="card p-0 overflow-hidden">
        <div style={{ maxHeight: "75vh", overflow: "auto" }}>
          <table
            className="text-sm border-collapse"
            style={{
              tableLayout: "fixed",
              width: ANCHO_TOTAL_BASE + (mesesVisibles ? totalSubcolumnasMeses * ANCHO_SUBCOL_MES : 0),
            }}
          >
            <colgroup>
              {COLUMNAS_BASE.map((c) => (
                <col key={c.clave} style={{ width: c.ancho }} />
              ))}
              {mesesVisibles &&
                mesesConSubcolumnas.flatMap((m) =>
                  m.subcolumnas.map((s) => <col key={`${m.mi}-${s}`} style={{ width: ANCHO_SUBCOL_MES }} />)
                )}
            </colgroup>
            <thead style={{ position: "sticky", top: 0, zIndex: 10, background: "var(--bg)" }}>
              <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--bg)" }}>
                <th
                  className="py-2 px-3 font-semibold align-top"
                  rowSpan={mesesVisibles ? 2 : 1}
                  style={{ ...estiloFijo("numero", "var(--bg)"), top: 0, zIndex: 11 }}
                >
                  N°
                </th>
                <th
                  className="py-2 px-3 font-semibold align-top"
                  rowSpan={mesesVisibles ? 2 : 1}
                  style={{ ...estiloFijo("empresa", "var(--bg)"), top: 0, zIndex: 11 }}
                >
                  Empresa
                </th>
                <th
                  className="py-2 px-3 font-semibold align-top"
                  rowSpan={mesesVisibles ? 2 : 1}
                  style={{ ...estiloFijo("grupo", "var(--bg)"), top: 0, zIndex: 11 }}
                >
                  Grupo de Gasto
                </th>
                <th
                  className="py-2 px-3 font-semibold align-top"
                  rowSpan={mesesVisibles ? 2 : 1}
                  style={{ ...estiloFijo("subgrupo", "var(--bg)"), top: 0, zIndex: 11 }}
                >
                  Subgrupo
                </th>
                <th
                  className="py-2 px-3 font-semibold align-top"
                  rowSpan={mesesVisibles ? 2 : 1}
                  style={{ ...estiloFijo("linea", "var(--bg)"), top: 0, zIndex: 11 }}
                >
                  Línea de Gasto
                </th>
                <th className="py-2 px-3 font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1}>
                  Responsable
                </th>
                <th className="py-2 px-3 font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1}>
                  Status
                </th>
                <th className="py-2 px-3 text-center font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1} style={ESTILO_DESTACADO}>
                  Presupuesto Aprobado
                </th>
                <th className="py-2 px-3 text-center font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1} style={ESTILO_DESTACADO}>
                  Gasto Real
                </th>
                <th className="py-2 px-3 text-center font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1} style={ESTILO_DESTACADO}>
                  Forecast
                </th>
                <th className="py-2 px-3 text-center font-semibold align-top" rowSpan={mesesVisibles ? 2 : 1} style={ESTILO_DESTACADO}>
                  Diferencia
                </th>
                {mesesVisibles &&
                  mesesConSubcolumnas.map((m) => (
                    <th
                      key={m.mi}
                      colSpan={m.subcolumnas.length}
                      className="py-1 px-3 text-center font-semibold whitespace-nowrap"
                      style={{ borderLeft: "1px solid var(--borde)" }}
                    >
                      {m.nombre}
                    </th>
                  ))}
              </tr>
              {mesesVisibles && (
                <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--bg)" }}>
                  {mesesConSubcolumnas.flatMap((m) =>
                    m.subcolumnas.map((s) => (
                      <th
                        key={`${m.mi}-${s}`}
                        className="py-1 px-3 text-center font-normal text-xs"
                        style={{ borderLeft: s === "real" ? "1px solid var(--borde)" : undefined }}
                      >
                        {s === "real" ? "Real" : "Proy."}
                      </th>
                    ))
                  )}
                </tr>
              )}
            </thead>
            <tbody>
              {filtradas.map((l, i) => {
                const resaltada = filaResaltada === l.filaExcel;
                return (
                <tr
                  key={l.filaExcel}
                  onClick={() => setFilaResaltada((f) => (f === l.filaExcel ? null : l.filaExcel))}
                  title="Clic para resaltar esta línea"
                  style={{
                    borderTop: "1px solid var(--borde)",
                    cursor: "pointer",
                    position: resaltada ? "relative" : undefined,
                    zIndex: resaltada ? 5 : undefined,
                    background: resaltada ? "#fff6d8" : undefined,
                    boxShadow: resaltada ? "inset 0 0 0 2px var(--alerta), 0 4px 14px rgba(0,0,0,0.18)" : undefined,
                    transform: resaltada ? "scale(1.01)" : undefined,
                    transition: "background 0.15s, box-shadow 0.15s, transform 0.15s",
                  }}
                >
                  <td
                    className="py-1.5 px-3 text-center"
                    style={{ ...estiloFijo("numero", resaltada ? "#fff6d8" : "var(--card)"), color: "var(--texto-suave)" }}
                  >
                    {i + 1}
                  </td>
                  <td
                    className="py-1.5 px-3 truncate"
                    style={{ ...estiloFijo("empresa", resaltada ? "#fff6d8" : "var(--card)"), color: "var(--texto-suave)" }}
                  >
                    {l.subNegocio}
                  </td>
                  <td className="py-1.5 px-3 truncate" style={estiloFijo("grupo", resaltada ? "#fff6d8" : "var(--card)")}>
                    {l.grupoNegocio}
                  </td>
                  <td className="py-1.5 px-3 truncate" style={estiloFijo("subgrupo", resaltada ? "#fff6d8" : "var(--card)")}>
                    {l.prioridad}
                  </td>
                  <td
                    className="py-1.5 px-3 truncate"
                    title={l.proyecto}
                    style={{ ...estiloFijo("linea", resaltada ? "#fff6d8" : "var(--card)"), maxWidth: 260 }}
                  >
                    {l.proyecto}
                  </td>
                  <td className="py-1.5 px-3">
                    <CampoEditable
                      fila={l.filaExcel}
                      campo="responsable"
                      tipo="texto"
                      valor={l.responsable}
                      endpoint="/api/opex/celda"
                      onGuardado={(v) => actualizarLocal(l.filaExcel, { responsable: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 px-3">
                    <select
                      className="campo text-xs font-medium"
                      style={{
                        padding: "0.25rem 0.5rem",
                        width: "100%",
                        maxWidth: "100%",
                        boxSizing: "border-box",
                        ...(l.status === "Activa"
                          ? { background: "#e3f3e3", color: "var(--exito)", borderColor: "var(--exito)" }
                          : l.status === "Inactiva"
                            ? { background: "#fde7e9", color: "var(--peligro)", borderColor: "var(--peligro)" }
                            : {}),
                      }}
                      value={["Activa", "Inactiva"].includes(l.status) ? l.status : ""}
                      onChange={(e) => guardarStatus(l.filaExcel, e.target.value)}
                    >
                      <option value="" disabled>
                        Sin definir
                      </option>
                      <option value="Activa">Activa</option>
                      <option value="Inactiva">Inactiva</option>
                    </select>
                  </td>
                  <td className="py-1.5 px-3 text-center text-xs whitespace-nowrap" style={ESTILO_DESTACADO}>
                    {moneda2(l.presupuestoAprobado)}
                    {mostrarSoles && (
                      <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                        {soles(l.presupuestoAprobado, tipoCambio)}
                      </span>
                    )}
                  </td>
                  <td className="py-1.5 px-3 text-center text-xs whitespace-nowrap" style={ESTILO_DESTACADO}>
                    {moneda2(l.gastoReal)}
                    {mostrarSoles && (
                      <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                        {soles(l.gastoReal, tipoCambio)}
                      </span>
                    )}
                  </td>
                  <td className="py-1.5 px-3 text-center text-xs whitespace-nowrap" style={ESTILO_DESTACADO}>
                    {moneda2(l.forecast)}
                    {mostrarSoles && (
                      <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                        {soles(l.forecast, tipoCambio)}
                      </span>
                    )}
                  </td>
                  <td
                    className="py-1.5 px-3 text-center text-xs whitespace-nowrap font-bold"
                    style={{ ...ESTILO_DESTACADO, color: l.diferencia < 0 ? "var(--peligro)" : "var(--exito)" }}
                  >
                    {moneda2(l.diferencia)}
                  </td>
                  {mesesVisibles &&
                    mesesConSubcolumnas.flatMap((m) =>
                      m.subcolumnas.map((s) => (
                        <td
                          key={`${m.mi}-${s}`}
                          className="py-1.5 px-3 overflow-hidden"
                          style={{ borderLeft: s === "real" ? "1px solid var(--borde)" : undefined }}
                        >
                          {s === "real" && m.mi < mesCierre ? (
                            <span
                              className="block text-center text-xs"
                              style={{ color: "var(--texto-suave)" }}
                              title="Mes cerrado — ya no se edita"
                            >
                              {moneda2(l.real[m.mi])}
                              {mostrarSoles && (
                                <span className="block text-[10px]" style={{ color: "var(--texto-suave)" }}>
                                  {soles(l.real[m.mi], tipoCambio)}
                                </span>
                              )}
                            </span>
                          ) : s === "real" ? (
                            <CampoMontoSumado
                              fila={l.filaExcel}
                              campo={`real:${m.mi}`}
                              valor={l.real[m.mi]}
                              className="text-center text-xs"
                              endpoint="/api/opex/celda"
                              mostrarSoles={mostrarSoles}
                              tipoCambio={tipoCambio}
                              onGuardado={(nuevo) => actualizarLocal(l.filaExcel, { real: l.real.map((x, j) => (j === m.mi ? nuevo : x)) })}
                            />
                          ) : (
                            <CampoMontoSumado
                              fila={l.filaExcel}
                              campo={`proyectado:${m.mi}`}
                              valor={l.proyectado[m.mi]}
                              className="text-center text-xs"
                              endpoint="/api/opex/celda"
                              mostrarSoles={mostrarSoles}
                              tipoCambio={tipoCambio}
                              onGuardado={(nuevo) =>
                                actualizarLocal(l.filaExcel, { proyectado: l.proyectado.map((x, j) => (j === m.mi ? nuevo : x)) })
                              }
                            />
                          )}
                        </td>
                      ))
                    )}
                </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr style={{ borderTop: "2px solid var(--borde)" }}>
                <td className="py-2 px-3 font-bold" colSpan={CANTIDAD_COLUMNAS_TEXTO} style={{ background: "var(--bg)" }}>
                  Total ({filtradas.length} líneas)
                </td>
                <td className="py-2 px-3 text-center text-xs font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.presupuestoAprobado)}
                  {mostrarSoles && (
                    <span className="block text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.presupuestoAprobado, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center text-xs font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.gastoReal)}
                  {mostrarSoles && (
                    <span className="block text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.gastoReal, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center text-xs font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.forecast)}
                  {mostrarSoles && (
                    <span className="block text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.forecast, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center text-xs font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.diferencia)}
                </td>
                {mesesVisibles && <td colSpan={totalSubcolumnasMeses} style={{ background: "var(--bg)" }}></td>}
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}

function FormularioNuevaLinea({
  gruposExistentes,
  onCancelar,
  onCreada,
}: {
  gruposExistentes: string[];
  onCancelar: () => void;
  onCreada: () => void;
}) {
  const [empresa, setEmpresa] = useState("");
  const [grupoGasto, setGrupoGasto] = useState("");
  const [grupoNuevo, setGrupoNuevo] = useState("");
  const [subgrupoGasto, setSubgrupoGasto] = useState("");
  const [lineaGasto, setLineaGasto] = useState("");
  const [moneda, setMoneda] = useState("USD");
  const [detalle, setDetalle] = useState("");
  const [responsable, setResponsable] = useState("");
  const [presupuestoAprobado, setPresupuestoAprobado] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const grupoFinal = grupoGasto === "__nuevo__" ? grupoNuevo.trim() : grupoGasto;
  const listo = empresa.trim() && grupoFinal && subgrupoGasto.trim() && lineaGasto.trim() && presupuestoAprobado.trim();

  async function crear() {
    setError(null);
    setEnviando(true);
    try {
      const res = await fetch("/api/opex/agregar-linea", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          empresa: empresa.trim(),
          grupoGasto: grupoFinal,
          subgrupoGasto: subgrupoGasto.trim(),
          lineaGasto: lineaGasto.trim(),
          moneda: moneda.trim() || "USD",
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
      <h2 className="font-semibold mb-3 text-sm">Nueva línea de gasto OPEX</h2>
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div>
          <span className="etiqueta mb-0">Empresa *</span>
          <input className="campo" value={empresa} onChange={(e) => setEmpresa(e.target.value)} disabled={enviando} />
        </div>
        <div>
          <span className="etiqueta mb-0">Grupo de Gasto *</span>
          <select className="campo" value={grupoGasto} onChange={(e) => setGrupoGasto(e.target.value)} disabled={enviando}>
            <option value="">Elige uno…</option>
            {gruposExistentes.map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
            <option value="__nuevo__">+ Grupo nuevo…</option>
          </select>
          {grupoGasto === "__nuevo__" && (
            <input
              className="campo mt-1"
              placeholder="Nombre del grupo nuevo"
              value={grupoNuevo}
              onChange={(e) => setGrupoNuevo(e.target.value)}
              disabled={enviando}
            />
          )}
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
      <p className="text-xs mt-3" style={{ color: "var(--texto-suave)" }}>
        Los 12 meses arrancan en 0 — se llenan después desde esta misma tabla o registrando facturas en Facturas OPEX. Status
        arranca en &quot;Activa&quot;.
      </p>
      {error && (
        <p className="text-sm mt-2" style={{ color: "var(--peligro)" }}>
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