"use client";

import { useEffect, useMemo, useState } from "react";
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
  COLORES_AVANCE,
  ETIQUETAS_AVANCE,
  NOMBRES_MES,
  NOMBRES_MES_CIERRE,
  ORDEN_AVANCE,
  avancePorGrupoNegocio,
  claveAvance,
  estaSuspendido,
  filtrarPorPrioridad,
  gastoPorGrupoNegocio,
  panoramaTrimestral,
  prioridadesDisponibles,
  resolverProyectos,
  type FilaTrimestral,
  type ItemConMeses,
  type PanoramaTrimestral,
  type ProyectoCapex,
  type ProyectoResuelto,
} from "@/lib/capex";
import { useMesCierre } from "@/lib/useMesCierre";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import { moneda, moneda2, monedaK, soles, solesK } from "@/lib/format";
import ControlTipoCambio from "@/components/ControlTipoCambio";
import MontoSoles from "@/components/MontoSoles";

// Anchos fijos de las tablas "Panorama" — mismo criterio que ya usa Presupuesto OPEX
// (table-layout:fixed + colgroup): sin esto, la tabla se estira a lo ancho de la card y
// los importes quedan con huecos grandes entre columnas en vez de alinearse compactos.
const ANCHO_COL_GRUPO_PANORAMA = 190;
const ANCHO_COL_TOTAL_PANORAMA = 100;
const ANCHO_COL_TRIMESTRE = 105;
const ANCHO_COL_MES_PANORAMA = 90;

// Paleta de las barras: tonos de azul (el mismo azul del template, en distintas
// intensidades) en vez de negro/gris — el rojo/rosa queda reservado para lo crítico
// (gasto por encima de lo planeado/aprobado), no para diferenciar categorías.
const COLOR_GRUPO: Record<string, string> = {
  EMISIVO: "#0f6cbd",
  RECEPTIVO: "#4a97d9",
  TRANSVERSAL: "#a9cdea",
};
/** Versión más clara de COLOR_GRUPO, para el segmento "Forecast" de las barras apiladas —
 *  mismo tono que el grupo, para poder diferenciar un grupo de otro de un vistazo. */
const COLOR_GRUPO_CLARO: Record<string, string> = {
  EMISIVO: "#4a97d9",
  RECEPTIVO: "#a9cdea",
  TRANSVERSAL: "#d4e6f4",
};

const COLORES_TRIMESTRE = ["#0f6cbd", "#4a97d9", "#8fc0ea", "#c3ddf3"];

/** Tonos tomados del template de diseño "Dashboard TI" compartido: azul para lo normal/
 *  informativo, rosa/magenta para lo crítico — en vez de rojo corporativo + gris neutro. */
const AZUL_TEMPLATE = "#0f6cbd";
const ROSA_TEMPLATE = "#d6246e";

/** Cartel en lenguaje simple junto al % de desviación — para no tener que interpretar el
 *  número solo por el color. */
// Cualquier valor positivo (por chico que sea) es gasto por encima de lo planeado — no
// hay "margen razonable" para eso, siempre se marca en rosa. Azul para todo lo demás.
function estadoDesviacion(pct: number): { texto: string; bg: string; color: string } {
  if (pct > 20) return { texto: "Muy por encima de lo planeado", bg: "#fbe1ec", color: ROSA_TEMPLATE };
  if (pct > 0) return { texto: "Por encima de lo planeado", bg: "#fbe1ec", color: ROSA_TEMPLATE };
  if (pct < 0) return { texto: "Por debajo de lo planeado", bg: "#e3edfa", color: AZUL_TEMPLATE };
  return { texto: "Alineado a lo planeado", bg: "#e3edfa", color: AZUL_TEMPLATE };
}

function CartelDesviacion({ pct }: { pct: number }) {
  const estado = estadoDesviacion(pct);
  return (
    <span
      className="inline-block rounded-full text-xs font-semibold px-2 py-0.5 whitespace-nowrap"
      style={{ background: estado.bg, color: estado.color }}
    >
      {estado.texto}
    </span>
  );
}

interface RespuestaCapex {
  proyectos: ProyectoCapex[];
  proyeccion: ItemConMeses[];
  archivo: string;
  actualizadoEn: string;
}

export default function DashboardCapex() {
  const [datos, setDatos] = useState<RespuestaCapex | null>(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [prioridadesSel, setPrioridadesSel] = usePersistedState<string[] | null>("capex-dashboard-prioridadesSel", null);
  const [mesCierre, setMesCierre] = useMesCierre();
  const [generandoCierre, setGenerandoCierre] = useState(false);
  const [mensajeCierre, setMensajeCierre] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("capex-dashboard-mostrarSoles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();

  function nombreSiguienteCierre(nombreActual: string): string | null {
    const m = nombreActual.match(/(\d+)\s*\+\s*(\d+)/);
    if (!m) return null;
    const cerrados = Number(m[1]);
    const restantes = Number(m[2]);
    if (restantes < 1) return null;
    return nombreActual.replace(m[0], `${cerrados + 1}+${restantes - 1}`);
  }

  async function generarArchivoDeCierre() {
    if (!datos) return;
    const nombreNuevo = nombreSiguienteCierre(datos.archivo);
    if (!nombreNuevo) {
      setMensajeCierre({ tipo: "error", texto: "No se pudo calcular el siguiente nombre de archivo." });
      return;
    }
    const confirmado = window.confirm(
      `Esto crea un archivo nuevo "${nombreNuevo}" en la misma carpeta de SharePoint (copia de "${datos.archivo}", con la fórmula de FORESCAST ajustada un mes). El archivo actual no se modifica. ¿Continuar?`
    );
    if (!confirmado) return;

    setGenerandoCierre(true);
    setMensajeCierre(null);
    try {
      const res = await fetch("/api/capex/cerrar-mes", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo generar el archivo.");
      setMensajeCierre({ tipo: "ok", texto: `Archivo "${json.archivo}" creado correctamente.` });
    } catch (e) {
      setMensajeCierre({ tipo: "error", texto: (e as Error).message });
    } finally {
      setGenerandoCierre(false);
    }
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/capex", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar el archivo.");
      setDatos(json);
      // Forma funcional (lee el estado más reciente, no el de cuando se llamó cargar): si
      // ya había una selección guardada (localStorage, restaurada async), no se pisa con
      // el default — solo aplica el default la primera vez que de verdad no hay nada.
      setPrioridadesSel((actual) => {
        if (actual !== null) return actual;
        const disponibles = prioridadesDisponibles(json.proyectos).map((p) => p.valor);
        const porDefecto = disponibles.filter((v) => v === "1" || v === "2");
        return porDefecto.length > 0 ? porDefecto : disponibles;
      });
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

  const proyectosCrudos = datos?.proyectos ?? [];
  const proyeccion = datos?.proyeccion ?? [];
  const prioridades = useMemo(() => prioridadesDisponibles(proyectosCrudos), [proyectosCrudos]);

  // Real + proyectado se combinan según el mes de cierre elegido — no según lo que diga
  // el Excel, así se puede mover sin tocar el archivo.
  const proyectos = useMemo(() => resolverProyectos(proyectosCrudos, mesCierre), [proyectosCrudos, mesCierre]);

  const filtrados = useMemo(
    () => (prioridadesSel ? filtrarPorPrioridad(proyectos, prioridadesSel) : proyectos),
    [proyectos, prioridadesSel]
  );
  // Un proyecto "Suspendido" sigue teniendo su Presupuesto Aprobado (esa plata ya se
  // asignó, no desaparece por pausarlo) — pero al no estar ejecutándose, no debe aportar
  // a Gasto Real ni Forecast en ningún total (KPIs, gráfico por grupo, Panorama actual,
  // sobrepasados). Por eso NO se filtran filas completas (eso también borraría su
  // Presupuesto Aprobado): se ponen en 0 solo Gasto Real/Forecast/meses de un Suspendido,
  // y su Diferencia queda en su Presupuesto Aprobado completo (todo sin usar, pausado).
  // "Status de proyectos" es la excepción a propósito: ese indicador sí necesita ver los
  // Suspendidos tal cual, porque es justamente donde se cuentan.
  const filtradosParaGasto = useMemo(
    () =>
      filtrados.map((p) =>
        estaSuspendido(p.status)
          ? {
              ...p,
              gastoReal: 0,
              forecast: 0,
              diferencia: p.presupuestoAprobado,
              meses: p.meses.map(() => 0),
              real: p.real.map(() => 0),
              proyectado: p.proyectado.map(() => 0),
            }
          : p
      ),
    [filtrados]
  );

  const porGrupo = useMemo(() => gastoPorGrupoNegocio(filtradosParaGasto), [filtradosParaGasto]);
  const panoramaActual = useMemo(() => panoramaTrimestral(filtradosParaGasto), [filtradosParaGasto]);
  // Línea base: solo Prioridad 1 y 2 — fija, no reacciona al filtro de prioridad de arriba
  // (que sí puede incluir 3, 4, etc.), es un recorte propio de este panorama.
  const proyeccionBase = useMemo(() => filtrarPorPrioridad(proyeccion, ["1", "2"]), [proyeccion]);
  const panoramaProyectado = useMemo(() => panoramaTrimestral(proyeccionBase), [proyeccionBase]);

  const totales = useMemo(
    () =>
      filtradosParaGasto.reduce(
        (acc, p) => ({
          presupuestoAprobado: acc.presupuestoAprobado + p.presupuestoAprobado,
          gastoReal: acc.gastoReal + p.gastoReal,
          forecast: acc.forecast + p.forecast,
          diferencia: acc.diferencia + p.diferencia,
        }),
        { presupuestoAprobado: 0, gastoReal: 0, forecast: 0, diferencia: 0 }
      ),
    [filtradosParaGasto]
  );

  function alternarPrioridad(valor: string) {
    setPrioridadesSel((prev) => {
      const actual = prev ?? [];
      return actual.includes(valor) ? actual.filter((v) => v !== valor) : [...actual, valor];
    });
  }

  function limpiarFiltros() {
    setPrioridadesSel(null);
  }

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
          <div className="flex flex-wrap gap-2 items-center">
            <span className="etiqueta mb-0">Prioridad:</span>
            {prioridades.map((p) => (
              <span
                key={p.valor}
                className="chip"
                data-activo={prioridadesSel?.includes(p.valor)}
                onClick={() => alternarPrioridad(p.valor)}
              >
                {p.valor} ({p.cantidad})
              </span>
            ))}
            {prioridadesSel && (
              <button
                type="button"
                className="text-xs font-semibold hover:underline"
                style={{ color: "var(--acento)" }}
                onClick={limpiarFiltros}
                title="Vuelve a mostrar todas las prioridades, sin ningún filtro de por medio"
              >
                Quitar filtros ✕
              </button>
            )}
          </div>
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
          <button
            className="boton-secundario"
            onClick={generarArchivoDeCierre}
            disabled={generandoCierre || !datos}
            title="Crea un archivo nuevo en SharePoint para el siguiente mes cerrado (ej. 7+5 → 8+4), igual que ya haces a mano cada mes."
          >
            {generandoCierre ? "Generando…" : "Generar archivo de cierre"}
          </button>
          <button className="boton-primario" onClick={cargar} disabled={cargando}>
            {cargando ? "Actualizando…" : "Actualizar"}
          </button>
        </div>
      </div>

      {mensajeCierre && (
        <div
          className="card p-3 text-sm"
          style={{
            color: mensajeCierre.tipo === "error" ? "var(--peligro)" : "var(--exito)",
            background: mensajeCierre.tipo === "error" ? undefined : "#e3f3e3",
          }}
        >
          {mensajeCierre.texto}
        </div>
      )}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <TarjetaKpi
          titulo="Presupuesto aprobado"
          valor={totales.presupuestoAprobado}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
        <TarjetaKpi
          titulo="Gasto real"
          valor={totales.gastoReal}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
        <TarjetaKpi
          titulo="Forecast (por ejecutar)"
          valor={totales.forecast}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
        <TarjetaKpi
          titulo="Diferencia"
          valor={totales.diferencia}
          color={totales.diferencia < 0 ? ROSA_TEMPLATE : AZUL_TEMPLATE}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
      </div>

      <div className="card p-4">
        <h2 className="font-semibold mb-4">Gasto por grupo de negocio</h2>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={porGrupo} layout="vertical" margin={{ left: 16 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--borde)" />
            <XAxis
              type="number"
              tickFormatter={(v) => (mostrarSoles ? solesK(v, tipoCambio) : moneda(v))}
              tick={{ fontSize: 11 }}
            />
            <YAxis type="category" dataKey="grupoNegocio" width={100} tick={{ fontSize: 12 }} />
            <Tooltip formatter={(v: number) => (mostrarSoles ? `${moneda(v)} · ${soles(v, tipoCambio)}` : moneda(v))} />
            <Legend />
            {/* El "fill" del Bar es solo lo que usa la leyenda como muestra de color — el
                color real de cada barra lo pone el <Cell> de abajo, por grupo. */}
            <Bar dataKey="gastoReal" name="Gasto real" stackId="a" fill={AZUL_TEMPLATE}>
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill={COLOR_GRUPO[g.grupoNegocio] ?? AZUL_TEMPLATE} />
              ))}
            </Bar>
            <Bar dataKey="forecast" name="Forecast" stackId="a" fill="#a9cdea">
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill={COLOR_GRUPO_CLARO[g.grupoNegocio] ?? "#a9cdea"} />
              ))}
            </Bar>
            <Bar dataKey="presupuestoAprobado" name="Presupuesto aprobado" fill="#c8c6c4">
              {porGrupo.map((g) => (
                <Cell key={g.grupoNegocio} fill="#201f1e" fillOpacity={0.15} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-4">
          {porGrupo.map((g) => (
            <div key={g.grupoNegocio} className="p-3 rounded-lg" style={{ background: "var(--bg)" }}>
              <p
                className="font-semibold text-sm mb-2"
                style={{ color: COLOR_GRUPO[g.grupoNegocio] ?? "var(--texto)", textTransform: "uppercase" }}
              >
                {g.grupoNegocio}
              </p>
              <table className="w-full text-xs">
                <tbody>
                  <tr>
                    <td className="py-0.5" style={{ color: "var(--texto-suave)" }}>
                      Presupuesto aprobado
                    </td>
                    <td className="py-0.5 text-right font-medium">
                      {moneda(g.presupuestoAprobado)}
                      <MontoSoles valorUsd={g.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} />
                    </td>
                  </tr>
                  <tr>
                    <td className="py-0.5" style={{ color: "var(--texto-suave)" }}>
                      − Gasto real
                    </td>
                    <td className="py-0.5 text-right font-medium">
                      {moneda(g.gastoReal)}
                      <MontoSoles valorUsd={g.gastoReal} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} />
                    </td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid var(--borde)" }}>
                    <td className="py-0.5" style={{ color: "var(--texto-suave)" }}>
                      − Forecast
                    </td>
                    <td className="py-0.5 text-right font-medium">
                      {moneda(g.forecast)}
                      <MontoSoles valorUsd={g.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} />
                    </td>
                  </tr>
                  <tr>
                    <td className="pt-1 font-semibold">= Diferencia</td>
                    <td
                      className="pt-1 text-right font-bold"
                      style={{ color: g.diferencia < 0 ? ROSA_TEMPLATE : AZUL_TEMPLATE }}
                    >
                      {moneda(g.diferencia)}
                      <MontoSoles valorUsd={g.diferencia} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} />
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          ))}
        </div>
      </div>

      <EstadoProyectosPorGrupoSection proyectos={filtrados} />

      {/* Lado a lado en pantallas grandes para comparar de un vistazo — cada tabla
          mantiene su propio scroll horizontal si no entra en la mitad de la pantalla. */}
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
        <PanoramaTrimestralSection
          titulo="Panorama proyectado 2026 (línea base)"
          panorama={panoramaProyectado}
          mesResaltado={mesCierre}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />

        <PanoramaTrimestralSection
          titulo="Panorama actual 2026"
          panorama={panoramaActual}
          mesCierre={mesCierre}
          mesResaltado={mesCierre}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
      </div>

      <PanoramaComparacionMatrizSection
        proyectado={panoramaProyectado}
        actual={panoramaActual}
        mostrarSoles={mostrarSoles}
        tipoCambio={tipoCambio}
      />
      <ProyectosSobrepasadosSection proyectos={filtradosParaGasto} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />
      <ProyectosCorridosSection
        proyeccionBase={proyeccionBase}
        proyectosCrudos={proyectosCrudos}
        mesCierre={mesCierre}
        mostrarSoles={mostrarSoles}
        tipoCambio={tipoCambio}
      />
    </div>
  );
}

/**
 * Cuántas líneas de BD_CAPEX hay en cada estado (No iniciado / Iniciado / Por culminar /
 * Culminado / Suspendido / Sin dato), por Grupo de Negocio — para ver de un vistazo, por
 * ejemplo, cuántos proyectos de EMISIVO todavía no arrancan o ya están cerrados. Mismos
 * buckets y colores que el selector de Avance en Detalle BD_CAPEX. Cada número es clic­able:
 * abre una ventana con el detalle de qué proyectos son.
 */
function EstadoProyectosPorGrupoSection({ proyectos }: { proyectos: ProyectoResuelto[] }) {
  const avancePorGrupo = useMemo(() => avancePorGrupoNegocio(proyectos), [proyectos]);
  const clavesConDatos = ORDEN_AVANCE.filter((clave) => avancePorGrupo.some((g) => (g.conteos[clave] ?? 0) > 0));
  const [celdaAbierta, setCeldaAbierta] = useState<{ grupo: string; clave: string } | null>(null);

  const proyectosDeLaCelda = useMemo(() => {
    if (!celdaAbierta) return [];
    return proyectos.filter(
      (p) => (p.grupoNegocio || "SIN GRUPO") === celdaAbierta.grupo && claveAvance(p) === celdaAbierta.clave
    );
  }, [proyectos, celdaAbierta]);

  if (avancePorGrupo.length === 0 || clavesConDatos.length === 0) return null;

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Status de proyectos</h2>
      <div className="overflow-x-auto">
        <table className="text-sm border-collapse" style={{ width: "100%" }}>
          <thead>
            <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
              <th className="py-2 pr-4 font-semibold">Grupo de negocio</th>
              {clavesConDatos.map((clave) => (
                <th key={clave} className="py-2 px-3 text-center font-semibold">
                  {ETIQUETAS_AVANCE[clave]}
                </th>
              ))}
              <th className="py-2 px-3 text-center font-semibold">Total</th>
            </tr>
          </thead>
          <tbody>
            {avancePorGrupo.map((g) => {
              const color = COLOR_GRUPO[g.grupoNegocio] ?? "var(--texto)";
              return (
                <tr key={g.grupoNegocio} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td className="py-1.5 pr-4 font-semibold" style={{ color, textTransform: "uppercase" }}>
                    {g.grupoNegocio}
                  </td>
                  {clavesConDatos.map((clave) => {
                    const cantidad = g.conteos[clave] ?? 0;
                    const estilo = COLORES_AVANCE[clave];
                    return (
                      <td key={clave} className="py-1.5 px-3 text-center">
                        {cantidad > 0 ? (
                          <button
                            type="button"
                            className="inline-block rounded-full text-xs font-semibold px-2 py-0.5 min-w-[28px] cursor-pointer"
                            style={{ background: estilo.bg, color: estilo.color, border: "none" }}
                            onClick={() => setCeldaAbierta({ grupo: g.grupoNegocio, clave })}
                            title="Toca para ver qué proyectos son"
                          >
                            {cantidad}
                          </button>
                        ) : (
                          <span style={{ color: "var(--texto-suave)" }}>—</span>
                        )}
                      </td>
                    );
                  })}
                  <td className="py-1.5 px-3 text-center font-bold">{g.total}</td>
                </tr>
              );
            })}
            <tr style={{ borderTop: "2px solid var(--borde)", background: "var(--acento-suave)" }}>
              <td className="py-2 pr-4 font-bold">Total general</td>
              {clavesConDatos.map((clave) => {
                const cantidad = avancePorGrupo.reduce((a, g) => a + (g.conteos[clave] ?? 0), 0);
                return (
                  <td key={clave} className="py-2 px-3 text-center font-bold">
                    {cantidad}
                  </td>
                );
              })}
              <td className="py-2 px-3 text-center font-bold">{avancePorGrupo.reduce((a, g) => a + g.total, 0)}</td>
            </tr>
          </tbody>
        </table>
      </div>

      {celdaAbierta && (
        <div
          className="fixed inset-0 flex items-center justify-center p-6"
          style={{ background: "rgba(0,0,0,0.4)", zIndex: 50 }}
          onClick={() => setCeldaAbierta(null)}
        >
          <div
            className="card p-4 w-full flex flex-col"
            style={{ maxWidth: 900, maxHeight: "85vh", background: "var(--card)" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-3 shrink-0">
              <h2 className="font-semibold">
                {celdaAbierta.grupo} — {ETIQUETAS_AVANCE[celdaAbierta.clave]} ({proyectosDeLaCelda.length})
              </h2>
              <button
                className="text-xl leading-none px-2"
                style={{ color: "var(--texto-suave)" }}
                onClick={() => setCeldaAbierta(null)}
              >
                ×
              </button>
            </div>
            <div style={{ overflow: "auto" }}>
              <table className="w-full text-sm border-collapse">
                <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                  <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--card)" }}>
                    <th className="py-2 pr-4 font-semibold">Proyecto</th>
                    <th className="py-2 pr-4 font-semibold">Detalle</th>
                    <th className="py-2 pr-4 font-semibold">Prioridad</th>
                    <th className="py-2 px-3 text-right font-semibold">Presupuesto Aprobado</th>
                  </tr>
                </thead>
                <tbody>
                  {proyectosDeLaCelda.map((p) => (
                    <tr key={p.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                      <td className="py-1.5 pr-4 font-semibold">{p.proyecto}</td>
                      <td className="py-1.5 pr-4" style={{ color: "var(--texto-suave)", maxWidth: 280 }} title={p.detalle}>
                        {p.detalle || "—"}
                      </td>
                      <td className="py-1.5 pr-4">{p.prioridad}</td>
                      <td className="py-1.5 px-3 text-right">{moneda2(p.presupuestoAprobado)}</td>
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

/**
 * Compara, por grupo de negocio y por trimestre a la vez, lo que se planificó al inicio
 * del año (línea base, Proyeccion 2026_vm) contra lo que se está gastando/ejecutando
 * ahora (BD_CAPEX) — para ver, por ejemplo, que TRANSVERSAL viene bien en T1-T3 pero se
 * dispara en T4, algo que un total por grupo o un total por trimestre por separado no
 * mostrarían.
 */
function PanoramaComparacionMatrizSection({
  proyectado,
  actual,
  mostrarSoles,
  tipoCambio,
}: {
  proyectado: PanoramaTrimestral;
  actual: PanoramaTrimestral;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const grupos = ["EMISIVO", "RECEPTIVO", "TRANSVERSAL"];
  const trimestres = ["T1", "T2", "T3", "T4"] as const;

  function celda(grupo: string, qi: number) {
    const proy = proyectado.grupos.find((g) => g.grupoNegocio === grupo)?.subtotal.t[qi] ?? 0;
    const act = actual.grupos.find((g) => g.grupoNegocio === grupo)?.subtotal.t[qi] ?? 0;
    const pct = proy > 0 ? ((act - proy) / proy) * 100 : 0;
    return { proy, act, pct, diferencia: act - proy };
  }

  const filas = grupos
    .map((grupo) => {
      const celdas = trimestres.map((_, qi) => celda(grupo, qi));
      const proyTotal = celdas.reduce((a, c) => a + c.proy, 0);
      const actTotal = celdas.reduce((a, c) => a + c.act, 0);
      const pctTotal = proyTotal > 0 ? ((actTotal - proyTotal) / proyTotal) * 100 : 0;
      return { grupo, celdas, proyTotal, actTotal, pctTotal };
    })
    .filter((f) => f.proyTotal > 0 || f.actTotal > 0);

  const totalesPorTrimestre = trimestres.map((_, qi) => {
    const proy = filas.reduce((a, f) => a + f.celdas[qi].proy, 0);
    const act = filas.reduce((a, f) => a + f.celdas[qi].act, 0);
    const pct = proy > 0 ? ((act - proy) / proy) * 100 : 0;
    return { proy, act, pct };
  });
  const proyGeneral = filas.reduce((a, f) => a + f.proyTotal, 0);
  const actGeneral = filas.reduce((a, f) => a + f.actTotal, 0);
  const pctGeneral = proyGeneral > 0 ? ((actGeneral - proyGeneral) / proyGeneral) * 100 : 0;

  function CeldaMatriz({ pct, diferencia }: { pct: number; diferencia: number }) {
    const estado = estadoDesviacion(pct);
    return (
      <td className="py-1.5 px-3 text-center" style={{ background: estado.bg }}>
        <div className="font-semibold" style={{ color: estado.color }}>
          {pct > 0 ? "+" : ""}
          {pct.toFixed(1)}%
        </div>
        <div className="text-[10px]" style={{ color: "var(--texto-suave)" }}>
          {diferencia > 0 ? "+" : ""}
          {moneda2(diferencia)}
        </div>
        {mostrarSoles && (
          <div className="text-[10px]" style={{ color: "var(--texto-suave)" }}>
            {diferencia > 0 ? "+" : ""}
            {soles(diferencia, tipoCambio)}
          </div>
        )}
      </td>
    );
  }

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Panorama actual vs. Panorama proyectado, por grupo y trimestre</h2>

      <table className="w-full text-sm border-collapse">
        <thead>
          <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
            <th className="py-2 pr-4 font-semibold">Grupo de negocio</th>
            {trimestres.map((t) => (
              <th key={t} className="py-2 px-3 text-center font-semibold">
                {t}
              </th>
            ))}
            <th className="py-2 px-3 text-center font-semibold">Total año</th>
          </tr>
        </thead>
        <tbody>
          {filas.map((f) => {
            const color = COLOR_GRUPO[f.grupo] ?? "var(--texto)";
            return (
              <tr key={f.grupo} style={{ borderTop: "1px solid var(--borde)" }}>
                <td className="py-1.5 pr-4 font-semibold" style={{ color, textTransform: "uppercase" }}>
                  {f.grupo}
                </td>
                {f.celdas.map((c, qi) => (
                  <CeldaMatriz key={qi} pct={c.pct} diferencia={c.diferencia} />
                ))}
                <CeldaMatriz pct={f.pctTotal} diferencia={f.actTotal - f.proyTotal} />
              </tr>
            );
          })}
          <tr style={{ borderTop: "2px solid var(--borde)", background: "var(--bg)" }}>
            <td className="py-1.5 pr-4 font-bold">Total general</td>
            {totalesPorTrimestre.map((t, qi) => (
              <CeldaMatriz key={qi} pct={t.pct} diferencia={t.act - t.proy} />
            ))}
            <CeldaMatriz pct={pctGeneral} diferencia={actGeneral - proyGeneral} />
          </tr>
        </tbody>
      </table>
    </div>
  );
}

const UMBRAL_MONTO_CORRIDO = 300; // ignora montos irrelevantes, no vale la pena alertar por eso
const UMBRAL_PCT_EJECUTADO_CORRIDO = 0.1; // menos del 10% ejecutado en un trimestre ya cerrado = alerta

/** Normaliza nombres de proyecto para cruzar la hoja de línea base con BD_CAPEX (mismo
 *  proyecto, dos hojas distintas) sin que un espacio de más o mayúsculas rompan el match. */
function normalizarNombreProyecto(nombre: string): string {
  return nombre.trim().toLowerCase().replace(/\s+/g, " ");
}

/**
 * Detecta líneas puntuales (Proyecto + Detalle) que tenían presupuesto proyectado en un
 * trimestre YA CERRADO (todos sus meses son "real", no forecast) pero casi no se ejecutó
 * nada — la señal de que esa línea se corrió a otro trimestre en vez de ejecutarse cuando
 * se planeó. Solo mira trimestres cerrados: en uno abierto, un real bajo puede ser
 * simplemente "todavía no llega ese mes", no un corrimiento real.
 *
 * Importante: "Proyecto" no es la unidad real de la hoja — cada proyecto trae varias
 * filas, una por Detalle (sub-tarea/línea de costo), y cada una puede tener su propia
 * Prioridad. Por eso el cruce contra BD_CAPEX es por Proyecto+Detalle, nunca agrupando
 * solo por nombre de proyecto (eso mezclaría líneas distintas del mismo proyecto).
 */
function ProyectosSobrepasadosSection({
  proyectos,
  mostrarSoles,
  tipoCambio,
}: {
  proyectos: ProyectoResuelto[];
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const [grupoAbierto, setGrupoAbierto] = useState<string | null>(null);

  const grupos = Array.from(new Set(proyectos.map((p) => p.grupoNegocio))).sort((a, b) => a.localeCompare(b, "es"));
  const porGrupo = grupos.map((grupo) => {
    const items = proyectos
      // -0.005: ignora diferencia "cero" con ruido de redondeo de floats (ej. -0.0000000001),
      // que de otro modo pasaría el filtro < 0 y se mostraría como sobrepasado en $0.00.
      .filter((p) => p.grupoNegocio === grupo && p.diferencia < -0.005)
      .sort((a, b) => a.diferencia - b.diferencia); // más sobrepasado primero (diferencia más negativa)
    const totalSobrepaso = items.reduce((a, p) => a + p.diferencia, 0);
    return { grupo, items, totalSobrepaso };
  });

  const grupoModal = porGrupo.find((g) => g.grupo === grupoAbierto);

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Proyectos que sobrepasaron su presupuesto, por grupo de negocio</h2>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        {porGrupo.map((g) => {
          const color = COLOR_GRUPO[g.grupo] ?? "var(--texto)";
          return (
            <button
              key={g.grupo}
              type="button"
              onClick={() => setGrupoAbierto(g.grupo)}
              disabled={g.items.length === 0}
              className="rounded-lg p-3 text-left"
              style={{
                background: "#fbe1ec",
                cursor: g.items.length === 0 ? "default" : "pointer",
                opacity: g.items.length === 0 ? 0.5 : 1,
              }}
              title={g.items.length > 0 ? "Toca para ver el detalle" : undefined}
            >
              <p className="text-2xl font-bold" style={{ color: ROSA_TEMPLATE }}>
                {g.items.length}
              </p>
              <p className="text-sm font-semibold" style={{ color }}>
                {g.grupo}
              </p>
            </button>
          );
        })}
      </div>

      {grupoModal && (
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
                {grupoModal.grupo} — proyectos sobrepasados ({grupoModal.items.length})
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
                  <th className="py-2 pr-4 font-semibold">Proyecto</th>
                  <th className="py-2 pr-4 font-semibold">Detalle</th>
                  <th className="py-2 px-3 text-right font-semibold">Presupuesto Aprobado</th>
                  <th className="py-2 px-3 text-right font-semibold">Gasto Real</th>
                  <th className="py-2 px-3 text-right font-semibold">Forecast</th>
                  <th className="py-2 px-3 text-right font-semibold">Diferencia</th>
                  <th className="py-2 px-3 text-right font-semibold">% Sobrepasado</th>
                </tr>
              </thead>
              <tbody>
                {grupoModal.items.map((p) => {
                  const sinPresupuesto = p.presupuestoAprobado <= 0;
                  const pctSobrepasado = sinPresupuesto ? null : (-p.diferencia / p.presupuestoAprobado) * 100;
                  return (
                    <tr key={p.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                      <td className="py-1.5 pr-4 font-semibold">{p.proyecto}</td>
                      <td
                        className="py-1.5 pr-4"
                        style={{ color: "var(--texto-suave)", maxWidth: 240 }}
                        title={p.detalle}
                      >
                        {p.detalle || "—"}
                      </td>
                      <td className="py-1.5 px-3 text-right">
                        {moneda2(p.presupuestoAprobado)}
                        <MontoSoles valorUsd={p.presupuestoAprobado} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs" />
                      </td>
                      <td className="py-1.5 px-3 text-right">
                        {moneda2(p.gastoReal)}
                        <MontoSoles valorUsd={p.gastoReal} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs" />
                      </td>
                      <td className="py-1.5 px-3 text-right">
                        {moneda2(p.forecast)}
                        <MontoSoles valorUsd={p.forecast} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs" />
                      </td>
                      <td className="py-1.5 px-3 text-right font-bold" style={{ color: ROSA_TEMPLATE }}>
                        {moneda2(p.diferencia)}
                        <MontoSoles valorUsd={p.diferencia} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
                      </td>
                      <td className="py-1.5 px-3 text-right font-medium" style={{ color: ROSA_TEMPLATE }}>
                        {sinPresupuesto ? "Sin presupuesto" : `${pctSobrepasado!.toFixed(0)}%`}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{ borderTop: "2px solid var(--borde)", background: "var(--acento-suave)" }}>
                  <td className="py-2 pr-4 font-bold" colSpan={2}>
                    Total
                  </td>
                  <td className="py-2 px-3 text-right font-bold">
                    {moneda2(grupoModal.items.reduce((a, p) => a + p.presupuestoAprobado, 0))}
                    <MontoSoles
                      valorUsd={grupoModal.items.reduce((a, p) => a + p.presupuestoAprobado, 0)}
                      tipoCambio={tipoCambio}
                      mostrarSoles={mostrarSoles}
                      className="block text-xs font-normal"
                    />
                  </td>
                  <td className="py-2 px-3 text-right font-bold">
                    {moneda2(grupoModal.items.reduce((a, p) => a + p.gastoReal, 0))}
                    <MontoSoles
                      valorUsd={grupoModal.items.reduce((a, p) => a + p.gastoReal, 0)}
                      tipoCambio={tipoCambio}
                      mostrarSoles={mostrarSoles}
                      className="block text-xs font-normal"
                    />
                  </td>
                  <td className="py-2 px-3 text-right font-bold">
                    {moneda2(grupoModal.items.reduce((a, p) => a + p.forecast, 0))}
                    <MontoSoles
                      valorUsd={grupoModal.items.reduce((a, p) => a + p.forecast, 0)}
                      tipoCambio={tipoCambio}
                      mostrarSoles={mostrarSoles}
                      className="block text-xs font-normal"
                    />
                  </td>
                  <td className="py-2 px-3 text-right font-bold" style={{ color: ROSA_TEMPLATE }}>
                    {moneda2(grupoModal.totalSobrepaso)}
                    <MontoSoles valorUsd={grupoModal.totalSobrepaso} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs font-normal" />
                  </td>
                  <td></td>
                </tr>
              </tfoot>
            </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ProyectosCorridosSection({
  proyeccionBase,
  proyectosCrudos,
  mesCierre,
  mostrarSoles,
  tipoCambio,
}: {
  proyeccionBase: ItemConMeses[];
  proyectosCrudos: ProyectoCapex[];
  mesCierre: number;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  type Alerta = {
    proyecto: string;
    detalle: string;
    grupo: string;
    prioridad: string;
    trimestre: string;
    proy: number;
    act: number;
    pctEjecutado: number;
    /** "corrido_confirmado": tiene proyección en otro trimestre Y ahí SÍ hay gasto real —
     *  confirmado que se ejecutó ahí en vez de cuando se planeó originalmente.
     *  "corrido_sin_confirmar": tiene proyección en OTRO TRIMESTRE YA CERRADO y tampoco
     *  se gastó nada ahí — evidencia real de que no se ejecutó en ninguno de los dos.
     *  "pendiente_futuro": la única proyección en otro lado cae en un trimestre que
     *  TODAVÍA NO LLEGA — no hay cómo saber si se va a ejecutar ahí o no, no es una
     *  alerta real todavía, solo algo a seguir mirando.
     *  "no_ejecutado": no hay proyección en ningún otro trimestre — no hay a dónde decir
     *  que se movió, simplemente no se ejecutó. */
    tipo: "corrido_confirmado" | "corrido_sin_confirmar" | "pendiente_futuro" | "no_ejecutado";
    trimestresDestino: string[];
    trimestresDestinoConfirmados: string[];
    trimestresDestinoFuturos: string[];
  };

  function clave(proyecto: string, detalle: string) {
    return `${normalizarNombreProyecto(proyecto)}|${normalizarNombreProyecto(detalle)}`;
  }

  // Suma defensiva por si BD_CAPEX también trae más de una fila para el mismo
  // Proyecto+Detalle — no debería, pero así no se pierde gasto real si pasa. También se
  // guarda si el proyecto quedó 100% Culminado (uno ya completado no "se corrió" solo
  // porque este trimestre no tenga gasto — probablemente ya gastó todo antes) y si está
  // Suspendido (uno pausado a propósito no ejecuta nada, no es una alerta real).
  const actualPorClave = new Map<string, { real: number[]; completado: boolean; suspendido: boolean }>();
  for (const p of proyectosCrudos) {
    if (!p.proyecto) continue;
    const k = clave(p.proyecto, p.detalle ?? "");
    const avance = parseFloat(String(p.avancePct).replace(",", ".")) || 0;
    const suspendido = estaSuspendido(p.status);
    const existente = actualPorClave.get(k);
    if (existente) {
      existente.real = existente.real.map((v, i) => v + (p.real[i] ?? 0));
      existente.completado = existente.completado || avance >= 1;
      existente.suspendido = existente.suspendido && suspendido;
    } else {
      actualPorClave.set(k, { real: [...p.real], completado: avance >= 1, suspendido });
    }
  }

  const alertas: Alerta[] = [];
  TRIMESTRES_MESES.forEach((meses, qi) => {
    const trimestreCerrado = meses[meses.length - 1] < mesCierre;
    if (!trimestreCerrado) return;
    for (const item of proyeccionBase) {
      if (!item.proyecto) continue;
      const proy = meses.reduce((a: number, mi) => a + (item.meses[mi] ?? 0), 0);
      if (proy < UMBRAL_MONTO_CORRIDO) continue;
      const entradaActual = actualPorClave.get(clave(item.proyecto, item.detalle ?? ""));
      if (entradaActual?.completado) continue; // ya 100% Culminado — no es un corrimiento
      if (entradaActual?.suspendido) continue; // pausado a propósito — no es una alerta real
      const realActual = entradaActual?.real;
      const act = realActual ? meses.reduce((a: number, mi) => a + (realActual[mi] ?? 0), 0) : 0;
      const pctEjecutado = proy > 0 ? act / proy : 0;
      if (pctEjecutado < UMBRAL_PCT_EJECUTADO_CORRIDO) {
        // ¿Esta misma línea tiene proyección en algún OTRO trimestre? Si sí, es candidato
        // a haberse corrido ahí — pero solo queda CONFIRMADO si en ese otro trimestre
        // también hay gasto real. Y si ese otro trimestre TODAVÍA NO CIERRA, no cuenta
        // como "tampoco se ejecutó ahí" — simplemente no le ha tocado su turno todavía.
        const destinos = TRIMESTRES_MESES.map((mesesOtro, qiOtro) => {
          if (qiOtro === qi) return null;
          const proyOtro = mesesOtro.reduce((a: number, mi) => a + (item.meses[mi] ?? 0), 0);
          if (proyOtro < UMBRAL_MONTO_CORRIDO) return null;
          const actOtro = realActual ? mesesOtro.reduce((a: number, mi) => a + (realActual[mi] ?? 0), 0) : 0;
          const cerradoOtro = mesesOtro[mesesOtro.length - 1] < mesCierre;
          return { trimestre: `T${qiOtro + 1}`, confirmado: actOtro > 0, cerrado: cerradoOtro };
        }).filter((d): d is { trimestre: string; confirmado: boolean; cerrado: boolean } => d !== null);

        const trimestresDestino = destinos.map((d) => d.trimestre);
        const trimestresDestinoConfirmados = destinos.filter((d) => d.confirmado).map((d) => d.trimestre);
        const destinosCerradosSinEjecutar = destinos.filter((d) => d.cerrado && !d.confirmado);
        const trimestresDestinoFuturos = destinos.filter((d) => !d.cerrado && !d.confirmado).map((d) => d.trimestre);

        alertas.push({
          proyecto: item.proyecto,
          detalle: item.detalle ?? "",
          grupo: item.grupoNegocio,
          prioridad: item.prioridad,
          trimestre: `T${qi + 1}`,
          proy,
          act,
          pctEjecutado,
          tipo:
            trimestresDestinoConfirmados.length > 0
              ? "corrido_confirmado"
              : destinosCerradosSinEjecutar.length > 0
                ? "corrido_sin_confirmar"
                : trimestresDestinoFuturos.length > 0
                  ? "pendiente_futuro"
                  : "no_ejecutado",
          trimestresDestino,
          trimestresDestinoConfirmados,
          trimestresDestinoFuturos,
        });
      }
    }
  });
  alertas.sort((a, b) => b.proy - a.proy);

  const [detalleTipo, setDetalleTipo] = useState<Alerta["tipo"] | "todos" | null>(null);
  const confirmados = alertas.filter((a) => a.tipo === "corrido_confirmado");
  const sinConfirmar = alertas.filter((a) => a.tipo === "corrido_sin_confirmar");
  const pendientes = alertas.filter((a) => a.tipo === "pendiente_futuro");
  const noEjecutados = alertas.filter((a) => a.tipo === "no_ejecutado");

  const TITULOS: Record<string, string> = {
    corrido_confirmado: "Se corrieron y se ejecutaron",
    corrido_sin_confirmar: "Con proyección en otro trimestre ya cerrado, sin ejecutar",
    pendiente_futuro: "Con proyección en un trimestre que todavía no llega",
    no_ejecutado: "No se ejecutaron en ningún trimestre",
    todos: "Todas las líneas",
  };
  const alertasDelModal = detalleTipo === "todos" ? alertas : alertas.filter((a) => a.tipo === detalleTipo);

  return (
    <div className="card p-4">
      <h2 className="font-semibold mb-4">Proyectos que parecen haberse corrido de trimestre</h2>

      {alertas.length === 0 ? (
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Sin alertas: no se encontraron proyectos con presupuesto planeado sin ejecutar en trimestres ya
          cerrados.
        </p>
      ) : (
        <>
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
            <ResumenAlerta
              cantidad={confirmados.length}
              etiqueta="Se corrieron y se ejecutaron"
              bg="#e3edfa"
              color={AZUL_TEMPLATE}
              onClick={() => setDetalleTipo("corrido_confirmado")}
            />
            <ResumenAlerta
              cantidad={sinConfirmar.length}
              etiqueta="En otro trimestre ya cerrado, sin ejecutar"
              bg="#fbe1ec"
              color={ROSA_TEMPLATE}
              onClick={() => setDetalleTipo("corrido_sin_confirmar")}
            />
            <ResumenAlerta
              cantidad={pendientes.length}
              etiqueta="En un trimestre que aún no llega — sin confirmar"
              bg="#f3f2f1"
              color="var(--texto-suave)"
              onClick={() => setDetalleTipo("pendiente_futuro")}
            />
            <ResumenAlerta
              cantidad={noEjecutados.length}
              etiqueta="No se ejecutaron en ningún trimestre"
              bg="#fbe1ec"
              color={ROSA_TEMPLATE}
              onClick={() => setDetalleTipo("no_ejecutado")}
            />
          </div>
          <button className="boton-secundario" onClick={() => setDetalleTipo("todos")}>
            Ver todas ({alertas.length} líneas)
          </button>
        </>
      )}

      {detalleTipo && (
        <ModalDetalleAlertas
          titulo={TITULOS[detalleTipo]}
          alertas={alertasDelModal}
          onCerrar={() => setDetalleTipo(null)}
          mostrarSoles={mostrarSoles}
          tipoCambio={tipoCambio}
        />
      )}
    </div>
  );
}

function ResumenAlerta({
  cantidad,
  etiqueta,
  bg,
  color,
  onClick,
}: {
  cantidad: number;
  etiqueta: string;
  bg: string;
  color: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={cantidad === 0}
      className="rounded-lg p-3 text-left"
      style={{ background: bg, cursor: cantidad === 0 ? "default" : "pointer", opacity: cantidad === 0 ? 0.5 : 1 }}
      title={cantidad > 0 ? "Toca para ver el detalle" : undefined}
    >
      <p className="text-2xl font-bold" style={{ color }}>
        {cantidad}
      </p>
      <p className="text-sm font-semibold" style={{ color }}>
        {etiqueta}
      </p>
    </button>
  );
}

function ModalDetalleAlertas({
  titulo,
  alertas,
  onCerrar,
  mostrarSoles,
  tipoCambio,
}: {
  titulo: string;
  mostrarSoles: boolean;
  tipoCambio: number;
  alertas: {
    proyecto: string;
    detalle: string;
    grupo: string;
    prioridad: string;
    trimestre: string;
    proy: number;
    act: number;
    pctEjecutado: number;
    tipo: "corrido_confirmado" | "corrido_sin_confirmar" | "pendiente_futuro" | "no_ejecutado";
    trimestresDestino: string[];
    trimestresDestinoConfirmados: string[];
    trimestresDestinoFuturos: string[];
  }[];
  onCerrar: () => void;
}) {
  return (
    <div
      className="fixed inset-0 flex items-center justify-center p-6"
      style={{ background: "rgba(0,0,0,0.4)", zIndex: 50 }}
      onClick={onCerrar}
    >
      <div
        className="card p-4 w-full flex flex-col"
        style={{ maxWidth: 1200, maxHeight: "85vh", background: "var(--card)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-3 shrink-0">
          <h2 className="font-semibold">
            {titulo} ({alertas.length})
          </h2>
          <button className="text-xl leading-none px-2" style={{ color: "var(--texto-suave)" }} onClick={onCerrar}>
            ×
          </button>
        </div>
        <div style={{ overflow: "auto" }}>
        <table className="w-full text-sm border-collapse">
          <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
            <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--card)" }}>
              <th className="py-2 pr-4 font-semibold">Proyecto</th>
              <th className="py-2 pr-4 font-semibold">Detalle</th>
              <th className="py-2 pr-4 font-semibold">Grupo</th>
              <th className="py-2 pr-4 font-semibold">Prioridad</th>
              <th className="py-2 pr-4 font-semibold">Trimestre</th>
              <th className="py-2 px-3 text-right font-semibold">Proyectado</th>
              <th className="py-2 px-3 text-right font-semibold">Ejecutado</th>
              <th className="py-2 px-3 text-right font-semibold">% Ejecutado</th>
              <th className="py-2 px-3 text-left font-semibold">Alerta</th>
            </tr>
          </thead>
          <tbody>
            {alertas.map((a, i) => {
              const color = COLOR_GRUPO[a.grupo] ?? "var(--texto)";
              return (
                <tr key={i} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td className="py-1.5 pr-4 font-semibold">{a.proyecto}</td>
                  <td className="py-1.5 pr-4" style={{ color: "var(--texto-suave)", maxWidth: 280 }} title={a.detalle}>
                    {a.detalle || "—"}
                  </td>
                  <td className="py-1.5 pr-4" style={{ color, textTransform: "uppercase" }}>
                    {a.grupo}
                  </td>
                  <td className="py-1.5 pr-4">{a.prioridad}</td>
                  <td className="py-1.5 pr-4 font-semibold">{a.trimestre}</td>
                  <td className="py-1.5 px-3 text-right">
                    {moneda2(a.proy)}
                    <MontoSoles valorUsd={a.proy} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs" />
                  </td>
                  <td className="py-1.5 px-3 text-right">
                    {moneda2(a.act)}
                    <MontoSoles valorUsd={a.act} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-xs" />
                  </td>
                  <td className="py-1.5 px-3 text-right font-medium" style={{ color: ROSA_TEMPLATE }}>
                    {(a.pctEjecutado * 100).toFixed(0)}%
                  </td>
                  <td className="py-1.5 px-3 text-left">
                    <span
                      className="inline-block rounded-full text-xs font-semibold px-2 py-0.5 whitespace-nowrap"
                      style={
                        a.tipo === "no_ejecutado" || a.tipo === "corrido_sin_confirmar"
                          ? { background: "#fbe1ec", color: ROSA_TEMPLATE }
                          : a.tipo === "pendiente_futuro"
                            ? { background: "#f3f2f1", color: "var(--texto-suave)" }
                            : { background: "#e3edfa", color: AZUL_TEMPLATE }
                      }
                    >
                      {a.tipo === "corrido_confirmado"
                        ? `Se corrió y se ejecutó en ${a.trimestresDestinoConfirmados.join(" y ")}`
                        : a.tipo === "corrido_sin_confirmar"
                          ? `Tiene proyección en ${a.trimestresDestino.join(" y ")} (ya cerrado), pero tampoco se ejecutó ahí`
                          : a.tipo === "pendiente_futuro"
                            ? `Tiene proyección en ${a.trimestresDestinoFuturos.join(" y ")}, todavía no llega ese trimestre`
                            : "No se ejecutó — sin proyección en otro trimestre"}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>
      </div>
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
  /** Muestra debajo el equivalente en Soles como referencia (no aplica a Diferencia). */
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
          S/{" "}
          {(valor * (tipoCambio ?? 1)).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </p>
      )}
    </div>
  );
}

interface ColumnaPanorama {
  clave: string;
  etiqueta: string;
  /** "Real" / "Proy." bajo el nombre del mes, solo cuando la columna viene de un
   *  trimestre desplegado — para dejar claro qué parte del total ya es gasto real. */
  subEtiqueta?: string;
  valor: (f: FilaTrimestral) => number;
  /** Este trimestre mezcla real+proyectado y está cerrado: mostrar el "+" para abrirlo. */
  expandible?: boolean;
  /** Primer mes de un trimestre ya abierto: mostrar el "−" para volver a cerrarlo. */
  colapsable?: boolean;
  quarterIndex?: number;
  /** "cerrado": ya pasó el mes de cierre completo (los 3 meses son "real"/ya ocurrió en
   *  el calendario). "parcial": el mes de cierre cae adentro (algunos meses ya pasaron,
   *  otros no). Se resalta para saber de un vistazo qué ya es historia y qué es plan. */
  estadoCierre?: "cerrado" | "parcial";
}

const TRIMESTRES_MESES = [
  [0, 1, 2],
  [3, 4, 5],
  [6, 7, 8],
  [9, 10, 11],
] as const;

/**
 * Arma las columnas de la tabla: un trimestre que todavía mezcla real con proyectado (mes
 * de cierre cae adentro de él) se muestra cerrado como un solo total con un "+" para
 * desplegarlo bajo pedido — mientras no se despliegue, no se distingue real de proyectado
 * en pantalla. Al desplegarlo se abre en sus 3 meses, cada uno etiquetado "Real" o "Proy."
 * según corresponda.
 *
 * `mesResaltado` es independiente de `permitirAbrir`: sirve solo para marcar qué
 * trimestres ya pasaron en el calendario (aunque esta tabla no tenga real/forecast propio,
 * como Panorama Proyectado, línea base fija).
 */
function construirColumnas(
  mesCierre: number | undefined,
  trimestresAbiertos: Set<number>,
  permitirAbrir: boolean,
  mesResaltado: number | undefined
): ColumnaPanorama[] {
  const columnas: ColumnaPanorama[] = [];
  TRIMESTRES_MESES.forEach((meses, qi) => {
    // Con mes de cierre (Panorama actual), cualquier trimestre se puede desplegar en sus
    // 3 meses bajo pedido — no solo el que mezcla real+proyectado (ej. T4, todo forecast,
    // también se abre igual para ver el detalle mes a mes).
    const puedeAbrirse = mesCierre != null && permitirAbrir;
    const abierto = puedeAbrirse && trimestresAbiertos.has(qi);

    let estadoCierre: "cerrado" | "parcial" | undefined;
    if (mesResaltado != null) {
      if (meses[meses.length - 1] < mesResaltado) estadoCierre = "cerrado";
      else if (meses[0] < mesResaltado) estadoCierre = "parcial";
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
          estadoCierre: mesResaltado != null ? (mi < mesResaltado ? "cerrado" : undefined) : undefined,
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

/** Fondo tenue para marcar de un vistazo qué trimestre/mes ya es historia (cerrado) o
 *  está a mitad de camino (parcial) — mismo criterio en Panorama Proyectado y Actual. */
function fondoCierre(estado: "cerrado" | "parcial" | undefined): string | undefined {
  if (estado === "cerrado") return "#e3edfa";
  if (estado === "parcial") return "#cfe0f3";
  return undefined;
}

function PanoramaTrimestralSection({
  titulo,
  panorama,
  mesCierre,
  permitirAbrir = true,
  mesResaltado,
  mostrarSoles,
  tipoCambio,
}: {
  titulo: string;
  panorama: PanoramaTrimestral;
  /** Si se pasa, el trimestre que todavía mezcla real+proyectado se abre en sus 3 meses. */
  mesCierre?: number;
  /** false en Panorama Proyectado: no tiene real/forecast propio, así que no se despliega
   *  mes a mes aunque se le pase mesResaltado solo para el resaltado de cerrados. */
  permitirAbrir?: boolean;
  /** Mes (1-12) desde donde ya pasó en el calendario — resalta T1/T2/T3... como cerrados. */
  mesResaltado?: number;
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const [trimestresAbiertos, setTrimestresAbiertos] = useState<Set<number>>(new Set());
  const columnas = useMemo(
    () => construirColumnas(mesCierre, trimestresAbiertos, permitirAbrir, mesResaltado),
    [mesCierre, trimestresAbiertos, permitirAbrir, mesResaltado]
  );

  function alternarTrimestre(qi: number) {
    setTrimestresAbiertos((prev) => {
      const siguiente = new Set(prev);
      if (siguiente.has(qi)) siguiente.delete(qi);
      else siguiente.add(qi);
      return siguiente;
    });
  }

  const datosChart = panorama.grupos.map((g) => ({
    grupo: g.grupoNegocio,
    T1: g.subtotal.t[0],
    T2: g.subtotal.t[1],
    T3: g.subtotal.t[2],
    T4: g.subtotal.t[3],
  }));

  return (
    <div className="card p-4 h-full flex flex-col">
      <h2 className="font-semibold mb-4">{titulo}</h2>

      <ResponsiveContainer width="100%" height={260}>
        <BarChart data={datosChart}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--borde)" />
          <XAxis dataKey="grupo" tick={{ fontSize: 12 }} />
          <YAxis
            tickFormatter={(v) => (mostrarSoles ? solesK(v, tipoCambio) : monedaK(v))}
            width={70}
            tick={{ fontSize: 11 }}
          />
          <Tooltip formatter={(v: number) => (mostrarSoles ? `${monedaK(v)} · ${solesK(v, tipoCambio)}` : monedaK(v))} />
          <Legend />
          {(["T1", "T2", "T3", "T4"] as const).map((t, i) => (
            <Bar key={t} dataKey={t} name={t} fill={COLORES_TRIMESTRE[i]} />
          ))}
        </BarChart>
      </ResponsiveContainer>

      <div className="overflow-x-auto mt-4">
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
              <th
                className="py-2 pr-4 font-semibold"
                style={{ position: "sticky", left: 0, zIndex: 2, background: "var(--card)", boxShadow: "2px 0 4px -2px rgba(0,0,0,0.15)" }}
              >
                Grupo / Prioridad
              </th>
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
                      c.expandible
                        ? "Ver el detalle mes a mes (real y proyectado)"
                        : c.colapsable
                          ? "Cerrar el detalle mensual"
                          : c.estadoCierre === "cerrado"
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
              <FragmentGrupo key={g.grupoNegocio} grupo={g} columnas={columnas} mostrarSoles={mostrarSoles} tipoCambio={tipoCambio} />
            ))}
            <tr style={{ background: "var(--acento-suave)", borderTop: "2px solid var(--borde)" }}>
              <td
                className="py-2 pr-4 font-bold"
                style={{ position: "sticky", left: 0, zIndex: 2, background: "var(--acento-suave)", boxShadow: "2px 0 4px -2px rgba(0,0,0,0.15)" }}
              >
                Total general
              </td>
              <td className="py-2 px-3 text-right font-bold">
                {moneda2(panorama.total.total)}
                <MontoSoles valorUsd={panorama.total.total} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
              </td>
              {columnas.map((c) => (
                <td
                  key={c.clave}
                  className="py-2 px-3 text-right font-bold"
                  style={{ background: fondoCierre(c.estadoCierre) }}
                >
                  {moneda2(c.valor(panorama.total))}
                  <MontoSoles valorUsd={c.valor(panorama.total)} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FragmentGrupo({
  grupo,
  columnas,
  mostrarSoles,
  tipoCambio,
}: {
  grupo: PanoramaTrimestral["grupos"][number];
  columnas: ColumnaPanorama[];
  mostrarSoles: boolean;
  tipoCambio: number;
}) {
  const color = COLOR_GRUPO[grupo.grupoNegocio] ?? "var(--texto)";
  return (
    <>
      {grupo.filas.map((fila) => (
        <tr key={fila.prioridad} style={{ borderTop: "1px solid var(--borde)" }}>
          <td
            className="py-1.5 pr-4 pl-4"
            style={{
              color: "var(--texto-suave)",
              textTransform: "uppercase",
              position: "sticky",
              left: 0,
              zIndex: 2,
              background: "var(--card)",
              boxShadow: "2px 0 4px -2px rgba(0,0,0,0.15)",
            }}
          >
            {grupo.grupoNegocio} — Prioridad {fila.prioridad}
          </td>
          <td className="py-1.5 px-3 text-right">
            {moneda2(fila.total)}
            <MontoSoles valorUsd={fila.total} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
          </td>
          {columnas.map((c) => (
            <td key={c.clave} className="py-1.5 px-3 text-right" style={{ background: fondoCierre(c.estadoCierre) }}>
              {moneda2(c.valor(fila))}
              <MontoSoles valorUsd={c.valor(fila)} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
            </td>
          ))}
        </tr>
      ))}
      <tr style={{ borderTop: "1px solid var(--borde)", background: "var(--bg)" }}>
        <td
          className="py-1.5 pr-4 font-semibold"
          style={{
            color,
            textTransform: "uppercase",
            position: "sticky",
            left: 0,
            zIndex: 2,
            background: "var(--bg)",
            boxShadow: "2px 0 4px -2px rgba(0,0,0,0.15)",
          }}
        >
          Total {grupo.grupoNegocio}
        </td>
        <td className="py-1.5 px-3 text-right font-semibold">
          {moneda2(grupo.subtotal.total)}
          <MontoSoles valorUsd={grupo.subtotal.total} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
        </td>
        {columnas.map((c) => (
          <td
            key={c.clave}
            className="py-1.5 px-3 text-right font-semibold"
            style={{ background: fondoCierre(c.estadoCierre) }}
          >
            {moneda2(c.valor(grupo.subtotal))}
            <MontoSoles valorUsd={c.valor(grupo.subtotal)} tipoCambio={tipoCambio} mostrarSoles={mostrarSoles} className="block text-[10px] font-normal" />
          </td>
        ))}
      </tr>
    </>
  );
}
