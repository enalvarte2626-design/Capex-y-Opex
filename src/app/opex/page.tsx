"use client";

import { Fragment, useEffect, useMemo, useState } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  NOMBRES_MES,
  NOMBRES_MES_CIERRE,
  gastoPorGrupoNegocio,
  mesesATrimestres,
  panoramaTrimestral,
  prioridadesDisponibles,
  resolverProyectos,
  type FilaTrimestral,
  type PanoramaTrimestral,
  type ProyectoCapex,
  type ProyectoResuelto,
} from "@/lib/capex";
import { useMesCierre } from "@/lib/useMesCierre";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import { moneda, moneda2, soles, solesK } from "@/lib/format";
import ControlTipoCambio from "@/components/ControlTipoCambio";
import FiltroMultiple from "@/components/FiltroMultiple";
import MontoSoles from "@/components/MontoSoles";

const COLORES = ["#c8102e", "#0f6cbd", "#107c10", "#d83b01", "#5c2d91", "#008272", "#986f0b", "#e3008c"];

/** Mismos tonos que CAPEX (template "Dashboard TI"): azul para lo normal/informativo,
 *  rosa/magenta para lo crítico — en vez de rojo/naranja/verde por separado. */
const AZUL_TEMPLATE = "#0f6cbd";
const ROSA_TEMPLATE = "#d6246e";

// Anchos fijos de las tablas "Panorama" — mismo criterio que ya usa Presupuesto OPEX
// (table-layout:fixed + colgroup): sin esto, la tabla se estira a lo ancho de la card y
// los importes quedan con huecos grandes entre columnas en vez de alinearse compactos.
const ANCHO_COL_GRUPO_PANORAMA = 190;
const ANCHO_COL_TOTAL_PANORAMA = 100;
const ANCHO_COL_TRIMESTRE = 105;
const ANCHO_COL_MES_PANORAMA = 90;

/** Etiqueta de estado según el % Ejecutado ((Gasto Real + Forecast) ÷ Presupuesto
 *  Aprobado — el año completo proyectado, no solo lo ya gastado, para que el signo
 *  siempre coincida con Diferencia) — misma idea que el badge de Avance en BD_CAPEX. */
function estadoEjecucion(pct: number): { texto: string; bg: string; color: string } {
  // 100.05: exactamente en el presupuesto (Diferencia = 0) puede llegar como 100.000001%
  // por ruido de redondeo de floats — eso no es "sobrepasado", es justo en el límite.
  if (pct > 150) return { texto: "Muy sobrepasado", bg: "#fbe1ec", color: ROSA_TEMPLATE };
  if (pct > 100.05) return { texto: "Sobrepasado", bg: "#fbe1ec", color: ROSA_TEMPLATE };
  if (pct >= 100) return { texto: "Ejecutado", bg: "#dceafa", color: AZUL_TEMPLATE };
  if (pct >= 80) return { texto: "Por agotarse", bg: "#dceafa", color: AZUL_TEMPLATE };
  return { texto: "En curso", bg: "#e3edfa", color: AZUL_TEMPLATE };
}

function EtiquetaEjecucion({ pct }: { pct: number }) {
  const estado = estadoEjecucion(pct);
  return (
    <span
      className="inline-block rounded-full text-xs font-semibold px-2 py-0.5 whitespace-nowrap"
      style={{ background: estado.bg, color: estado.color }}
    >
      {estado.texto}
    </span>
  );
}

interface RespuestaOpex {
  lineas: ProyectoCapex[];
  archivo: string;
  actualizadoEn: string;
}

export default function DashboardOpex() {
  const [datos, setDatos] = useState<RespuestaOpex | null>(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [gruposSel, setGruposSel] = usePersistedState<string[] | null>("opex-dashboard-gruposSel", null);
  const [mesCierre, setMesCierre] = useMesCierre();
  const [mostrarSoles, setMostrarSoles] = usePersistedState("opex-dashboard-mostrarSoles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/opex", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setDatos(json);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const lineasCrudas = datos?.lineas ?? [];
  const grupos = useMemo(() => prioridadesDisponibles(lineasCrudas.map((l) => ({ prioridad: l.grupoNegocio }))), [
    lineasCrudas,
  ]);

  const lineas = useMemo(() => resolverProyectos(lineasCrudas, mesCierre), [lineasCrudas, mesCierre]);
  const filtradas = useMemo(
    () => (gruposSel ? lineas.filter((l) => gruposSel.includes(l.grupoNegocio)) : lineas),
    [lineas, gruposSel]
  );

  const porGrupo = useMemo(() => gastoPorGrupoNegocio(filtradas), [filtradas]);
  const panorama = useMemo(() => panoramaTrimestral(filtradas), [filtradas]);

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

  const colorPorGrupo = useMemo(() => {
    const mapa: Record<string, string> = {};
    grupos.forEach((g, i) => {
      mapa[g.valor] = COLORES[i % COLORES.length];
    });
    return mapa;
  }, [grupos]);

  if (cargando && !datos) {
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
    <div className="flex flex-col gap-6">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex flex-wrap gap-4 items-center">
          <FiltroMultiple
            etiqueta="Grupo de Gasto"
            opciones={grupos.map((g) => ({ valor: g.valor, cantidad: g.cantidad }))}
            seleccion={gruposSel ?? grupos.map((g) => g.valor)}
            onCambiar={setGruposSel}
          />
          {gruposSel && (
            <button
              type="button"
              className="text-xs font-semibold hover:underline"
              style={{ color: "var(--acento)" }}
              onClick={() => setGruposSel(null)}
              title="Vuelve a mostrar todos los grupos de gasto, sin ningún filtro de por medio"
            >
              Quitar filtros ✕
            </button>
          )}
          <div className="flex items-center gap-2">
            <span className="etiqueta mb-0">Mes de cierre:</span>
            <select
              className="campo"
              style={{ width: "auto" }}
              value={mesCierre}
              onChange={(e) => setMesCierre(Number(e.target.value))}
              title="Último mes con Gasto Real cerrado — de ahí en adelante se cuenta como Forecast. No cambia el Excel, solo cómo se ve aquí."
            >
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  {nombre}
                </option>
              ))}
            </select>
          </div>
          <span
            className="chip"
            data-activo={mostrarSoles}
            onClick={() => setMostrarSoles((v) => !v)}
            title="Muestra el equivalente en Soles como referencia — no afecta ningún cálculo ni se guarda en el Excel."
          >
            {mostrarSoles ? "✓ " : ""}Habilitar en Soles
          </span>
          {mostrarSoles && <ControlTipoCambio tipoCambio={tipoCambio} onCambiar={setTipoCambio} />}
        </div>
        <div className="flex items-center gap-3">
          {datos && (
            <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
              Actualizado {new Date(datos.actualizadoEn).toLocaleString("es-PE")}
            </span>
          )}
          <button className="boton-primario" onClick={cargar} disabled={cargando}>
            {cargando ? "Actualizando…" : "Actualizar"}
          </button>
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <TarjetaKpi titulo="Presupuesto aprobado" valor={totales.presupuestoAprobado} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />
        <TarjetaKpi titulo="Gasto real" valor={totales.gastoReal} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />
        <TarjetaKpi titulo="Forecast (por ejecutar)" valor={totales.forecast} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />
        <TarjetaKpi
          titulo="Diferencia"
          valor={totales.diferencia}
          color={totales.diferencia < 0 ? ROSA_TEMPLATE : AZUL_TEMPLATE}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
      </div>

      <div className="card p-4">
        <h2 className="font-semibold mb-4">Gasto por grupo de gasto</h2>
        <ResponsiveContainer width="100%" height={Math.max(220, porGrupo.length * 34)}>
          <BarChart data={porGrupo} layout="vertical" margin={{ left: 16 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--borde)" />
            <XAxis
              type="number"
              tickFormatter={(v) => (mostrarSoles ? solesK(v, tipoCambio) : moneda(v))}
              tick={{ fontSize: 11 }}
            />
            <YAxis type="category" dataKey="grupoNegocio" width={140} tick={{ fontSize: 11 }} />
            <Tooltip formatter={(v: number) => (mostrarSoles ? `${moneda(v)} · ${soles(v, tipoCambio)}` : moneda(v))} />
            <Legend />
            {/* fill del <Bar> = lo que muestra la Leyenda (un color fijo por serie); el
                color real de cada barra lo pone el <Cell> de abajo, por grupo. */}
            <Bar dataKey="gastoReal" name="Gasto real" stackId="a" fill={AZUL_TEMPLATE}>
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill={colorPorGrupo[g.grupoNegocio] ?? AZUL_TEMPLATE} />
              ))}
            </Bar>
            <Bar dataKey="forecast" name="Forecast" stackId="a" fill="#a9cdea" fillOpacity={0.45}>
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill={colorPorGrupo[g.grupoNegocio] ?? AZUL_TEMPLATE} />
              ))}
            </Bar>
            <Bar dataKey="presupuestoAprobado" name="Presupuesto aprobado" fill="#c8c6c4">
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill="#201f1e" fillOpacity={0.15} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <TablaLineasOpex filtradas={filtradas} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />

      <LineasSinGastoSection filtradas={filtradas} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />

      <PanoramaTrimestralSection
        titulo="Panorama OPEX 2026"
        panorama={panorama}
        lineas={filtradas}
        mesCierre={mesCierre}
        mostrarSoles={mostrarSoles}
        tipoCambio={tipoCambio}
      />
    </div>
  );
}

function TarjetaKpi({
  titulo,
  valor,
  color,
  mostrarSoles,
  tipoCambio,
}: {
  titulo: string;
  valor: number;
  color?: string;
  mostrarSoles?: boolean;
  tipoCambio?: number;
}) {
  return (
    <div className="card p-4">
      <p className="etiqueta">{titulo}</p>
      <p className="kpi-valor" style={{ color: color ?? "var(--texto)" }}>
        {moneda2(valor)}
      </p>
      {mostrarSoles && (
        <p className="text-sm mt-0.5" style={{ color: "var(--texto-suave)" }}>
          S/ {(valor * (tipoCambio ?? 1)).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </p>
      )}
    </div>
  );
}

/**
 * Cuadro con todas las líneas de gasto OPEX (Presupuesto/Real/Forecast/Diferencia +
 * % ejecutado), y debajo un gráfico de cuáles consumen más — ambos respetan el filtro
 * de Grupo de Gasto de arriba, así que excluir un grupo ahí también los actualiza aquí.
 */
interface ResumenGrupo {
  grupo: string;
  presupuestoAprobado: number;
  gastoReal: number;
  forecast: number;
  diferencia: number;
  cantidadLineas: number;
}

function TablaLineasOpex({
  filtradas,
  mostrarSoles,
  tipoCambio,
}: {
  filtradas: ReturnType<typeof resolverProyectos>;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const porGrupo = useMemo(() => {
    const mapa = new Map<string, ResumenGrupo>();
    for (const l of filtradas) {
      const actual = mapa.get(l.grupoNegocio) ?? {
        grupo: l.grupoNegocio,
        presupuestoAprobado: 0,
        gastoReal: 0,
        forecast: 0,
        diferencia: 0,
        cantidadLineas: 0,
      };
      actual.presupuestoAprobado += l.presupuestoAprobado;
      actual.gastoReal += l.gastoReal;
      actual.forecast += l.forecast;
      actual.diferencia += l.diferencia;
      actual.cantidadLineas += 1;
      mapa.set(l.grupoNegocio, actual);
    }
    return Array.from(mapa.values()).sort((a, b) => b.gastoReal - a.gastoReal);
  }, [filtradas]);

  const totalGeneral = useMemo(
    () =>
      porGrupo.reduce(
        (acc, g) => ({
          presupuestoAprobado: acc.presupuestoAprobado + g.presupuestoAprobado,
          gastoReal: acc.gastoReal + g.gastoReal,
          forecast: acc.forecast + g.forecast,
          diferencia: acc.diferencia + g.diferencia,
        }),
        { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 }
      ),
    [porGrupo]
  );

  const [grupoAbierto, setGrupoAbierto] = useState<string | null>(null);
  // Solo las líneas que realmente aportan a la diferencia negativa del grupo — una línea
  // en positivo o sin gasto no explica el sobregasto, solo agregaría ruido a la ventana.
  const lineasDelGrupo = useMemo(
    () =>
      grupoAbierto
        ? [...filtradas]
            .filter((l) => l.grupoNegocio === grupoAbierto && l.diferencia < -0.005)
            .sort((a, b) => a.diferencia - b.diferencia)
        : [],
    [filtradas, grupoAbierto]
  );

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Todas las líneas OPEX — resumen por grupo</h2>
      <table className="w-full text-sm border-collapse">
        <thead>
          <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
            <th className="py-2 pr-4 font-semibold">Grupo de Gasto</th>
            <th className="py-2 px-3 text-right font-semibold">Presupuesto</th>
            <th className="py-2 px-3 text-right font-semibold">Gasto Real</th>
            <th className="py-2 px-3 text-right font-semibold">Forecast</th>
            <th className="py-2 px-3 text-right font-semibold">Diferencia</th>
            <th className="py-2 px-3 text-right font-semibold">% Ejecutado</th>
          </tr>
        </thead>
        <tbody>
          {porGrupo.map((g) => {
            // Real + Forecast (no solo Real) — así este % siempre coincide con el signo
            // de Diferencia, que también compara contra el año completo proyectado.
            const pctEjecutado = g.presupuestoAprobado > 0 ? ((g.gastoReal + g.forecast) / g.presupuestoAprobado) * 100 : 0;
            const sobrepasado = g.diferencia < -0.005; // ignora diferencia "cero" con ruido de redondeo de floats
            return (
              <tr
                key={g.grupo}
                style={{
                  borderTop: "1px solid var(--borde)",
                  cursor: sobrepasado ? "pointer" : undefined,
                }}
                onClick={sobrepasado ? () => setGrupoAbierto(g.grupo) : undefined}
                title={sobrepasado ? "Toca para ver qué líneas de este grupo están generando la diferencia" : undefined}
              >
                <td className="py-1.5 pr-4">
                  {g.grupo}
                  <span className="ml-1 text-xs" style={{ color: "var(--texto-suave)" }}>
                    ({g.cantidadLineas})
                  </span>
                  {sobrepasado && (
                    <span className="ml-1 font-bold" style={{ color: ROSA_TEMPLATE }}>
                      ↗
                    </span>
                  )}
                </td>
                <td className="py-1.5 px-3 text-right whitespace-nowrap">
                  {moneda2(g.presupuestoAprobado)}
                  <MontoSoles valorUsd={g.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                </td>
                <td className="py-1.5 px-3 text-right whitespace-nowrap">
                  {moneda2(g.gastoReal)}
                  <MontoSoles valorUsd={g.gastoReal} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                </td>
                <td className="py-1.5 px-3 text-right whitespace-nowrap">
                  {moneda2(g.forecast)}
                  <MontoSoles valorUsd={g.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                </td>
                <td
                  className="py-1.5 px-3 text-right whitespace-nowrap font-medium"
                  style={{ color: g.diferencia < 0 ? ROSA_TEMPLATE : AZUL_TEMPLATE }}
                >
                  {moneda2(g.diferencia)}
                  <MontoSoles valorUsd={g.diferencia} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                </td>
                <td className="py-1.5 px-3 text-right">
                  <div className="font-medium">{pctEjecutado.toFixed(0)}%</div>
                  <EtiquetaEjecucion pct={pctEjecutado} />
                </td>
              </tr>
            );
          })}
        </tbody>
        <tfoot>
          <tr style={{ borderTop: "2px solid var(--borde)", background: "var(--acento-suave)" }}>
            <td className="py-2 pr-4 font-bold">Total general</td>
            <td className="py-2 px-3 text-right font-bold">
              {moneda2(totalGeneral.presupuestoAprobado)}
              <MontoSoles valorUsd={totalGeneral.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
            </td>
            <td className="py-2 px-3 text-right font-bold">
              {moneda2(totalGeneral.gastoReal)}
              <MontoSoles valorUsd={totalGeneral.gastoReal} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
            </td>
            <td className="py-2 px-3 text-right font-bold">
              {moneda2(totalGeneral.forecast)}
              <MontoSoles valorUsd={totalGeneral.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
            </td>
            <td className="py-2 px-3 text-right font-bold">
              {moneda2(totalGeneral.diferencia)}
              <MontoSoles valorUsd={totalGeneral.diferencia} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
            </td>
            <td className="py-2 px-3 text-right">
              {(() => {
                const pctTotal =
                  totalGeneral.presupuestoAprobado > 0
                    ? ((totalGeneral.gastoReal + totalGeneral.forecast) / totalGeneral.presupuestoAprobado) * 100
                    : 0;
                return (
                  <>
                    <div className="font-bold">{pctTotal.toFixed(0)}%</div>
                    <EtiquetaEjecucion pct={pctTotal} />
                  </>
                );
              })()}
            </td>
          </tr>
        </tfoot>
      </table>

      {grupoAbierto && (
        <div
          className="fixed inset-0 flex items-center justify-center p-6"
          style={{ background: "rgba(0,0,0,0.4)", zIndex: 50 }}
          onClick={() => setGrupoAbierto(null)}
        >
          <div
            className="card p-4 w-full flex flex-col"
            style={{ maxWidth: 1000, maxHeight: "85vh", background: "var(--card)" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-3 shrink-0">
              <h2 className="font-semibold">
                {grupoAbierto} — líneas que generan la diferencia ({lineasDelGrupo.length})
              </h2>
              <button
                className="text-xl leading-none px-2"
                style={{ color: "var(--texto-suave)" }}
                onClick={() => setGrupoAbierto(null)}
              >
                ×
              </button>
            </div>
            <div style={{ overflow: "auto" }}>
            <table className="w-full text-sm border-collapse">
              <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--card)" }}>
                  <th className="py-2 pr-4 font-semibold">Línea de Gasto</th>
                  <th className="py-2 px-3 text-right font-semibold">Presupuesto</th>
                  <th className="py-2 px-3 text-right font-semibold">Gasto Real</th>
                  <th className="py-2 px-3 text-right font-semibold">Forecast</th>
                  <th className="py-2 px-3 text-right font-semibold">Diferencia</th>
                </tr>
              </thead>
              <tbody>
                {lineasDelGrupo.map((l) => (
                  <tr key={l.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                    <td className="py-1.5 pr-4 font-medium">{l.proyecto}</td>
                    <td className="py-1.5 px-3 text-right whitespace-nowrap">
                      {moneda2(l.presupuestoAprobado)}
                      <MontoSoles valorUsd={l.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                    </td>
                    <td className="py-1.5 px-3 text-right whitespace-nowrap">
                      {moneda2(l.gastoReal)}
                      <MontoSoles valorUsd={l.gastoReal} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                    </td>
                    <td className="py-1.5 px-3 text-right whitespace-nowrap">
                      {moneda2(l.forecast)}
                      <MontoSoles valorUsd={l.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                    </td>
                    <td
                      className="py-1.5 px-3 text-right whitespace-nowrap font-bold"
                      style={{ color: l.diferencia < 0 ? ROSA_TEMPLATE : AZUL_TEMPLATE }}
                    >
                      {moneda2(l.diferencia)}
                      <MontoSoles valorUsd={l.diferencia} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/** Grupos que no aportan a la alerta "sin gasto ni proyección": accesorios/piezas menores
 *  cuya compra es puntual y no planeada con anticipación — que no tengan forecast cargado
 *  no es una señal de alerta para este grupo. */
const GRUPOS_EXCLUIDOS_ALERTA_SIN_GASTO = ["PARTES Y PIEZAS"];

/** Líneas puntuales que no aportan a la alerta porque su gasto real ya se registra bajo
 *  otra línea del Excel (ej. Copilot Studio se factura junto con Office 365) — así que un
 *  Gasto Real en $0 aquí no significa que no se haya ejecutado nada. */
const LINEAS_EXCLUIDAS_ALERTA_SIN_GASTO = ["Copilot Studio (Ex power agents)"];

/**
 * Alerta real para OPEX: líneas presupuestales sin NADA de gasto — ni ejecutado (Gasto
 * Real) ni proyectado para lo que resta del año (Forecast). No basta con que no se haya
 * gastado todavía si ya hay una proyección cargada para más adelante; la alerta es para
 * las que ni siquiera tienen eso: nadie planeó ni ejecutó nada para esa línea.
 */
function LineasSinGastoSection({
  filtradas,
  mostrarSoles,
  tipoCambio,
}: {
  filtradas: ReturnType<typeof resolverProyectos>;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const sinGasto = useMemo(
    () =>
      [...filtradas]
        // < medio centavo: ignora ruido de redondeo, no solo === 0. Solo interesa la
        // alerta más grave: se aprobó presupuesto y no se ejecutó ni se proyectó nada —
        // una línea que nunca tuvo presupuesto no es una alerta real.
        .filter(
          (l) =>
            l.gastoReal < 0.005 &&
            l.forecast < 0.005 &&
            l.presupuestoAprobado >= 0.005 &&
            !GRUPOS_EXCLUIDOS_ALERTA_SIN_GASTO.includes(l.grupoNegocio) &&
            !LINEAS_EXCLUIDAS_ALERTA_SIN_GASTO.includes(l.proyecto)
        )
        .sort((a, b) => a.grupoNegocio.localeCompare(b.grupoNegocio, "es") || a.proyecto.localeCompare(b.proyecto, "es")),
    [filtradas]
  );
  const [abierto, setAbierto] = useState(false);

  if (filtradas.length === 0) return null;

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Líneas con presupuesto aprobado sin ejecutar</h2>
      <button
        type="button"
        onClick={() => sinGasto.length > 0 && setAbierto(true)}
        disabled={sinGasto.length === 0}
        className="rounded-lg p-3 text-left"
        style={{
          background: sinGasto.length > 0 ? "#fbe1ec" : "var(--bg)",
          cursor: sinGasto.length > 0 ? "pointer" : "default",
        }}
        title={sinGasto.length > 0 ? "Toca para ver el detalle" : undefined}
      >
        <p className="text-2xl font-bold" style={{ color: sinGasto.length > 0 ? ROSA_TEMPLATE : "var(--texto-suave)" }}>
          {sinGasto.length}
        </p>
        <p className="text-sm font-semibold" style={{ color: "var(--texto-suave)" }}>
          de {filtradas.length} líneas, según el filtro de arriba
        </p>
      </button>

      {abierto && (
        <div
          className="fixed inset-0 flex items-center justify-center p-6"
          style={{ background: "rgba(0,0,0,0.4)", zIndex: 50 }}
          onClick={() => setAbierto(false)}
        >
          <div
            className="card p-4 w-full flex flex-col"
            style={{ maxWidth: 1000, maxHeight: "85vh", background: "var(--card)" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-3 shrink-0">
              <h2 className="font-semibold">Líneas con presupuesto aprobado sin ejecutar ({sinGasto.length})</h2>
              <button
                className="text-xl leading-none px-2"
                style={{ color: "var(--texto-suave)" }}
                onClick={() => setAbierto(false)}
              >
                ×
              </button>
            </div>
            <div style={{ overflow: "auto" }}>
              <table className="w-full text-sm border-collapse">
                <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                  <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--card)" }}>
                    <th className="py-2 pr-4 font-semibold">Grupo de Gasto</th>
                    <th className="py-2 pr-4 font-semibold">Línea</th>
                    <th className="py-2 pr-4 font-semibold">Detalle</th>
                    <th className="py-2 px-3 text-right font-semibold">Presupuesto Aprobado</th>
                    <th className="py-2 px-3 text-right font-semibold">Forecast</th>
                  </tr>
                </thead>
                <tbody>
                  {sinGasto.map((l) => (
                    <tr key={l.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                      <td className="py-1.5 pr-4">{l.grupoNegocio}</td>
                      <td className="py-1.5 pr-4 font-medium">{l.proyecto}</td>
                      <td className="py-1.5 pr-4" style={{ color: "var(--texto-suave)", maxWidth: 240 }} title={l.detalle}>
                        {l.detalle || "—"}
                      </td>
                      <td className="py-1.5 px-3 text-right whitespace-nowrap font-bold" style={{ color: ROSA_TEMPLATE }}>
                        {moneda2(l.presupuestoAprobado)}
                        <MontoSoles valorUsd={l.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                      </td>
                      <td className="py-1.5 px-3 text-right whitespace-nowrap">
                        {moneda2(l.forecast)}
                        <MontoSoles valorUsd={l.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="text-xs" />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

interface ColumnaPanorama {
  clave: string;
  etiqueta: string;
  subEtiqueta?: string;
  valor: (f: FilaTrimestral) => number;
  expandible?: boolean;
  colapsable?: boolean;
  quarterIndex?: number;
  /** "cerrado": ya pasó el mes de cierre completo (los 3 meses son "real"/ya ocurrió en
   *  el calendario). "parcial": el mes de cierre cae adentro (algunos meses ya pasaron,
   *  otros no). Se resalta para saber de un vistazo qué ya es historia y qué es plan —
   *  mismo criterio que Panorama Actual en CAPEX. */
  estadoCierre?: "cerrado" | "parcial";
}

const TRIMESTRES_MESES = [
  [0, 1, 2],
  [3, 4, 5],
  [6, 7, 8],
  [9, 10, 11],
] as const;

function construirColumnas(mesCierre: number | undefined, trimestresAbiertos: Set<number>): ColumnaPanorama[] {
  const columnas: ColumnaPanorama[] = [];
  TRIMESTRES_MESES.forEach((meses, qi) => {
    const puedeAbrirse = mesCierre != null;
    const abierto = puedeAbrirse && trimestresAbiertos.has(qi);

    let estadoCierre: "cerrado" | "parcial" | undefined;
    if (mesCierre != null) {
      if (meses[meses.length - 1] < mesCierre) estadoCierre = "cerrado";
      else if (meses[0] < mesCierre) estadoCierre = "parcial";
    }

    if (abierto) {
      meses.forEach((mi, idx) => {
        columnas.push({
          clave: `m${mi}`,
          etiqueta: NOMBRES_MES[mi],
          subEtiqueta: mi < mesCierre! ? "Real" : "Proy.",
          valor: (f) => f.meses[mi] ?? 0,
          colapsable: idx === 0,
          quarterIndex: qi,
          estadoCierre: mi < mesCierre! ? "cerrado" : undefined,
        });
      });
    } else {
      columnas.push({
        clave: `t${qi}`,
        etiqueta: `T${qi + 1}`,
        valor: (f) => f.t[qi],
        expandible: puedeAbrirse,
        quarterIndex: qi,
        estadoCierre,
      });
    }
  });
  return columnas;
}

/** Fondo tenue para marcar de un vistazo qué trimestre/mes ya es historia (cerrado) o está
 *  a mitad de camino (parcial) — mismo criterio que Panorama Actual en CAPEX. */
function fondoCierre(estado: "cerrado" | "parcial" | undefined): string | undefined {
  if (estado === "cerrado") return "#e3edfa";
  if (estado === "parcial") return "#cfe0f3";
  return undefined;
}

/** Convierte una línea presupuestal (ya resuelta con Real+Forecast por mes) al mismo
 *  formato FilaTrimestral que usan los subtotales — así las columnas T1-T4 / mes a mes ya
 *  armadas (construirColumnas) sirven igual para una línea puntual que para un subtotal. */
function filaTrimestralDeLinea(l: ProyectoResuelto): FilaTrimestral {
  const t = mesesATrimestres(l.meses);
  return { prioridad: l.prioridad, t, total: t.reduce((a, b) => a + b, 0), meses: l.meses };
}

function PanoramaTrimestralSection({
  titulo,
  panorama,
  lineas,
  mesCierre,
  mostrarSoles,
  tipoCambio,
}: {
  titulo: string;
  panorama: PanoramaTrimestral;
  /** Líneas crudas (mismo filtro ya aplicado arriba) — para poder listar el detalle por
   *  línea presupuestal al expandir un grupo, no solo el subtotal por prioridad. */
  lineas: ProyectoResuelto[];
  mesCierre?: number;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const [trimestresAbiertos, setTrimestresAbiertos] = useState<Set<number>>(new Set());
  const [gruposAbiertos, setGruposAbiertos] = useState<Set<string>>(new Set());
  const columnas = useMemo(() => construirColumnas(mesCierre, trimestresAbiertos), [mesCierre, trimestresAbiertos]);

  const lineasPorGrupo = useMemo(() => {
    const mapa = new Map<string, ProyectoResuelto[]>();
    for (const l of lineas) {
      const clave = l.grupoNegocio || "SIN GRUPO";
      if (!mapa.has(clave)) mapa.set(clave, []);
      mapa.get(clave)!.push(l);
    }
    return mapa;
  }, [lineas]);

  function alternarTrimestre(qi: number) {
    setTrimestresAbiertos((prev) => {
      const siguiente = new Set(prev);
      if (siguiente.has(qi)) siguiente.delete(qi);
      else siguiente.add(qi);
      return siguiente;
    });
  }

  function alternarGrupo(grupo: string) {
    setGruposAbiertos((prev) => {
      const siguiente = new Set(prev);
      if (siguiente.has(grupo)) siguiente.delete(grupo);
      else siguiente.add(grupo);
      return siguiente;
    });
  }

  function celda(v: number) {
    return mostrarSoles ? (
      <>
        {moneda2(v)}
        <div className="text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
          S/ {(v * tipoCambio).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </div>
      </>
    ) : (
      moneda2(v)
    );
  }

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">{titulo}</h2>
      <div className="overflow-x-auto">
        <table
          className="text-xs border-collapse"
          style={{
            tableLayout: "fixed",
            width:
              ANCHO_COL_GRUPO_PANORAMA +
              ANCHO_COL_TOTAL_PANORAMA +
              columnas.reduce((a, c) => a + (c.subEtiqueta ? ANCHO_COL_MES_PANORAMA : ANCHO_COL_TRIMESTRE), 0),
          }}
        >
          <colgroup>
            <col style={{ width: ANCHO_COL_GRUPO_PANORAMA }} />
            <col style={{ width: ANCHO_COL_TOTAL_PANORAMA }} />
            {columnas.map((c) => (
              <col key={c.clave} style={{ width: c.subEtiqueta ? ANCHO_COL_MES_PANORAMA : ANCHO_COL_TRIMESTRE }} />
            ))}
          </colgroup>
          <thead>
            <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
              <th className="py-2 pr-4 font-semibold">Grupo / Subgrupo</th>
              <th className="py-2 px-3 text-right font-semibold">Total</th>
              {columnas.map((c) => {
                const accionable = c.expandible || c.colapsable;
                return (
                  <th
                    key={c.clave}
                    className="py-2 px-3 text-right font-semibold"
                    style={{
                      ...(accionable ? { cursor: "pointer", userSelect: "none" } : undefined),
                      background: fondoCierre(c.estadoCierre),
                    }}
                    onClick={accionable ? () => alternarTrimestre(c.quarterIndex!) : undefined}
                    title={
                      c.estadoCierre === "cerrado"
                        ? "Trimestre ya cerrado"
                        : c.estadoCierre === "parcial"
                          ? "Trimestre en curso — mezcla meses cerrados y por venir"
                          : undefined
                    }
                  >
                    {c.etiqueta}
                    {c.expandible && (
                      <span className="ml-1 font-bold" style={{ color: "var(--acento)" }}>
                        +
                      </span>
                    )}
                    {c.colapsable && (
                      <span className="ml-1 font-bold" style={{ color: "var(--acento)" }}>
                        −
                      </span>
                    )}
                    {c.subEtiqueta && (
                      <div className="text-[10px] font-normal" style={{ color: "var(--texto-suave)" }}>
                        {c.subEtiqueta}
                      </div>
                    )}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {panorama.grupos.map((g) => (
              <FragmentGrupo
                key={g.grupoNegocio}
                grupo={g}
                lineasDelGrupo={lineasPorGrupo.get(g.grupoNegocio) ?? []}
                columnas={columnas}
                celda={celda}
                abierto={gruposAbiertos.has(g.grupoNegocio)}
                onAlternar={() => alternarGrupo(g.grupoNegocio)}
              />
            ))}
            <tr style={{ background: "var(--acento-suave)", borderTop: "2px solid var(--borde)" }}>
              <td className="py-2 pr-4 font-bold">Total general</td>
              <td className="py-2 px-3 text-right font-bold">{celda(panorama.total.total)}</td>
              {columnas.map((c) => (
                <td key={c.clave} className="py-2 px-3 text-right font-bold" style={{ background: fondoCierre(c.estadoCierre) }}>
                  {celda(c.valor(panorama.total))}
                </td>
              ))}
            </tr>
            <tr style={{ borderBottom: "1px solid var(--borde)" }}>
              <td className="py-1.5 pr-4 font-semibold" style={{ color: "var(--texto-suave)" }}>
                % del total del año
              </td>
              <td className="py-1.5 px-3 text-right font-semibold" style={{ color: "var(--texto-suave)" }}>
                100%
              </td>
              {columnas.map((c) => {
                const pct = panorama.total.total > 0 ? (c.valor(panorama.total) / panorama.total.total) * 100 : 0;
                return (
                  <td
                    key={c.clave}
                    className="py-1.5 px-3 text-right font-semibold"
                    style={{ color: "var(--texto-suave)", background: fondoCierre(c.estadoCierre) }}
                  >
                    {pct.toFixed(1)}%
                  </td>
                );
              })}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FragmentGrupo({
  grupo,
  lineasDelGrupo,
  columnas,
  celda,
  abierto,
  onAlternar,
}: {
  grupo: PanoramaTrimestral["grupos"][number];
  lineasDelGrupo: ProyectoResuelto[];
  columnas: ColumnaPanorama[];
  celda: (v: number) => React.ReactNode;
  abierto: boolean;
  onAlternar: () => void;
}) {
  const lineasPorPrioridad = useMemo(() => {
    const mapa = new Map<string, ProyectoResuelto[]>();
    for (const l of lineasDelGrupo) {
      const clave = l.prioridad || "Sin prioridad";
      if (!mapa.has(clave)) mapa.set(clave, []);
      mapa.get(clave)!.push(l);
    }
    return mapa;
  }, [lineasDelGrupo]);

  // Las 3 líneas de mayor consumo del grupo (Real+Forecast del año) — se marcan con 🔥
  // para ubicarlas de un vistazo entre todas las líneas del grupo, sin tener que leer
  // cada monto una por una.
  const topConsumoIds = useMemo(() => {
    return new Set(
      lineasDelGrupo
        .map((l) => ({ filaExcel: l.filaExcel, total: filaTrimestralDeLinea(l).total }))
        .filter((x) => x.total > 0)
        .sort((a, b) => b.total - a.total)
        .slice(0, 3)
        .map((x) => x.filaExcel)
    );
  }, [lineasDelGrupo]);

  return (
    <>
      <tr style={{ borderTop: "1px solid var(--borde)", background: "var(--bg)" }}>
        <td
          className="py-1.5 pr-4 font-semibold"
          style={{ cursor: "pointer", userSelect: "none" }}
          onClick={onAlternar}
          title={abierto ? "Cerrar el detalle por línea presupuestal" : "Ver el detalle por línea presupuestal"}
        >
          {grupo.grupoNegocio}
          <span className="ml-1 font-bold" style={{ color: "var(--acento)" }}>
            {abierto ? "−" : "+"}
          </span>
        </td>
        <td className="py-1.5 px-3 text-right font-semibold">{celda(grupo.subtotal.total)}</td>
        {columnas.map((c) => (
          <td key={c.clave} className="py-1.5 px-3 text-right font-semibold" style={{ background: fondoCierre(c.estadoCierre) }}>
            {celda(c.valor(grupo.subtotal))}
          </td>
        ))}
      </tr>
      {abierto &&
        grupo.filas.map((fila) => (
          <Fragment key={fila.prioridad}>
            <tr style={{ borderTop: "1px solid var(--borde)" }}>
              <td className="py-1.5 pr-4 pl-6" style={{ color: "var(--texto-suave)" }}>
                {fila.prioridad}
              </td>
              <td className="py-1.5 px-3 text-right">{celda(fila.total)}</td>
              {columnas.map((c) => (
                <td key={c.clave} className="py-1.5 px-3 text-right" style={{ background: fondoCierre(c.estadoCierre) }}>
                  {celda(c.valor(fila))}
                </td>
              ))}
            </tr>
            {(lineasPorPrioridad.get(fila.prioridad || "Sin prioridad") ?? []).map((l) => {
              const filaLinea = filaTrimestralDeLinea(l);
              const esTopConsumo = topConsumoIds.has(l.filaExcel);
              return (
                <tr key={l.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td
                    className="py-1 pr-4 pl-10 text-[11px]"
                    style={{ color: "var(--texto-suave)" }}
                    title={l.detalle || l.proyecto}
                  >
                    {esTopConsumo && (
                      <span className="mr-1" title="Entre las líneas de mayor consumo del grupo">
                        🔥
                      </span>
                    )}
                    {l.proyecto}
                    {l.detalle && <span> — {l.detalle}</span>}
                  </td>
                  <td className="py-1 px-3 text-right text-[11px]">{celda(filaLinea.total)}</td>
                  {columnas.map((c) => (
                    <td key={c.clave} className="py-1 px-3 text-right text-[11px]" style={{ background: fondoCierre(c.estadoCierre) }}>
                      {celda(c.valor(filaLinea))}
                    </td>
                  ))}
                </tr>
              );
            })}
          </Fragment>
        ))}
    </>
  );
}