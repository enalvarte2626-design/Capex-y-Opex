"use client";

import { useEffect, useMemo, useState } from "react";
import {
  ETIQUETAS_AVANCE,
  MARCA_SUSPENDIDO,
  NOMBRES_MES_CIERRE,
  ORDEN_AVANCE,
  claveAvance,
  estaSuspendido,
  prioridadesDisponibles,
  resolverProyectos,
  valorAvance,
  type ProyectoCapex,
  type ProyectoResuelto,
} from "@/lib/capex";
import { moneda2 } from "@/lib/format";
import { useMesCierre } from "@/lib/useMesCierre";
import { useTipoCambio } from "@/lib/useTipoCambio";
import { usePersistedState } from "@/lib/usePersistedState";
import CampoEditable from "@/components/CampoEditable";
import CampoMontoSumado from "@/components/CampoMontoSumado";
import ControlTipoCambio from "@/components/ControlTipoCambio";

const NOMBRES_MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

/** Minúsculas y sin tildes, para que buscar "optimizacion" encuentre "Optimización". */
const RANGO_DIACRITICOS = /[\u0300-\u036f]/g;

function normalizar(texto: string): string {
  return texto.toLowerCase().normalize("NFD").replace(RANGO_DIACRITICOS, "");
}

type Columna =
  | "proyecto"
  | "detalle"
  | "subNegocio"
  | "grupoNegocio"
  | "prioridad"
  | "avancePct"
  | "presupuestoAprobado"
  | "gastoReal"
  | "forecast"
  | "diferencia";

function soles(valorUsd: number, tipoCambio: number): string {
  return `S/ ${(valorUsd * tipoCambio).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

const ANCHO_SUBCOL_MES = 90;
const CANTIDAD_COLUMNAS_TEXTO = 7; // # + proyecto..avance, antes de los 4 importes destacados
const ANCHO_NUMERO = 44;
const COLOR_RESALTADO = "#fff3b0"; // amarillo, alto contraste para exponer en pantalla

const COLUMNAS_BASE: { clave: Columna; etiqueta: string; numerica?: boolean; destacada?: boolean; ancho: number }[] = [
  { clave: "proyecto", etiqueta: "Proyecto", ancho: 170 },
  { clave: "detalle", etiqueta: "Detalle", ancho: 170 },
  { clave: "subNegocio", etiqueta: "Sub. Negocio", ancho: 100 },
  { clave: "grupoNegocio", etiqueta: "Grupo", ancho: 90 },
  { clave: "prioridad", etiqueta: "Prioridad", ancho: 95 },
  { clave: "avancePct", etiqueta: "Avance", ancho: 150 },
  { clave: "presupuestoAprobado", etiqueta: "Presupuesto aprobado", numerica: true, destacada: true, ancho: 115 },
  { clave: "gastoReal", etiqueta: "Gasto real", numerica: true, destacada: true, ancho: 100 },
  { clave: "forecast", etiqueta: "Forecast", numerica: true, destacada: true, ancho: 100 },
  { clave: "diferencia", etiqueta: "Diferencia", numerica: true, destacada: true, ancho: 100 },
];

// Ancho SIEMPRE fijo en px (nunca porcentual): con anchos relativos al contenedor, en
// pantallas angostas los importes quedaban recortados detrás de la columna siguiente. Con
// px fijos el texto nunca pierde su espacio — si no entra en la pantalla, aparece scroll
// horizontal en vez de esconder dígitos.
const ANCHO_TOTAL_BASE = ANCHO_NUMERO + COLUMNAS_BASE.reduce((a, c) => a + c.ancho, 0);
const COLUMNAS = COLUMNAS_BASE;

/** Estilo de las 4 columnas clave (Presupuesto/Real/Forecast/Diferencia): resaltadas del resto. */
const ESTILO_DESTACADO = { background: "var(--acento-suave)" };

function estadoAvance(avancePct: string, status: string): { texto: string; bg: string; color: string } {
  if (estaSuspendido(status)) {
    return { texto: "Suspendido", bg: "#f3f2f1", color: "var(--peligro)" };
  }
  const valor = valorAvance(avancePct);
  switch (valor) {
    case 1:
      return { texto: "100% · Culminado", bg: "#e3f3e3", color: "var(--exito)" };
    case 0.8:
      return { texto: "80% · Por culminar", bg: "#fdecdc", color: "var(--alerta)" };
    case 0.3:
      return { texto: "30% · Iniciado", bg: "#e3edfa", color: "#0f6cbd" };
    case 0:
      return { texto: "0% · No iniciado", bg: "var(--bg)", color: "var(--texto-suave)" };
    default:
      return { texto: "—", bg: "var(--bg)", color: "var(--texto-suave)" };
  }
}

const OPCIONES_AVANCE: { valor: string; texto: string }[] = [
  { valor: "0", texto: "0% · No iniciado" },
  { valor: "0.3", texto: "30% · Iniciado" },
  { valor: "0.8", texto: "80% · Por culminar" },
  { valor: "1", texto: "100% · Culminado" },
  { valor: "suspendido", texto: "Suspendido" },
];

interface OpcionFiltro {
  clave: string;
  texto: string;
  cantidad: number;
}

/**
 * Desplegable de filtro que vive dentro del propio <th> de la columna — mismo criterio para
 * Grupo, Prioridad, Avance, Proyecto y Detalle: nada de botones sueltos arriba de la tabla,
 * el filtro está pegado a la columna que filtra (como una columna con filtro en Excel).
 */
function FiltroColumna({
  abierto,
  onAbrir,
  onCerrar,
  opciones,
  seleccionados,
  onCambiarSeleccion,
  conBusqueda,
}: {
  abierto: boolean;
  onAbrir: () => void;
  onCerrar: () => void;
  opciones: OpcionFiltro[];
  seleccionados: string[] | null;
  onCambiarSeleccion: (v: string[] | null) => void;
  conBusqueda?: boolean;
}) {
  const [busqueda, setBusqueda] = useState("");

  useEffect(() => {
    if (!abierto) setBusqueda("");
  }, [abierto]);

  const visibles =
    conBusqueda && busqueda.trim()
      ? opciones.filter((o) => normalizar(o.texto).includes(normalizar(busqueda)))
      : opciones;

  function alternarOpcion(valor: string) {
    const actual = seleccionados ?? [];
    const nueva = actual.includes(valor) ? actual.filter((v) => v !== valor) : [...actual, valor];
    onCambiarSeleccion(nueva.length === 0 ? null : nueva);
  }

  return (
    <>
      <span
        className="cursor-pointer ml-1"
        style={{ color: seleccionados ? "var(--acento-fuerte)" : "var(--texto-suave)" }}
        title="Filtrar"
        onClick={(e) => {
          e.stopPropagation();
          abierto ? onCerrar() : onAbrir();
        }}
      >
        ▾{seleccionados ? ` (${seleccionados.length})` : ""}
      </span>
      {abierto && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 15 }} onClick={onCerrar} />
          <div
            className="card p-2"
            style={{
              position: "absolute",
              top: "100%",
              left: 0,
              zIndex: 20,
              minWidth: 210,
              maxHeight: 320,
              overflow: "auto",
              background: "var(--card)",
              boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {conBusqueda && (
              <input
                type="text"
                className="campo text-xs mb-1"
                placeholder="Buscar…"
                value={busqueda}
                onChange={(e) => setBusqueda(e.target.value)}
                autoFocus
              />
            )}
            {visibles.length === 0 && (
              <p className="text-xs py-1" style={{ color: "var(--texto-suave)" }}>
                Sin resultados
              </p>
            )}
            {visibles.map((o) => (
              <label
                key={o.clave}
                className="flex items-center gap-2 text-xs font-normal py-1 cursor-pointer whitespace-nowrap"
                style={{ color: "var(--texto)" }}
                title={o.texto}
              >
                <input type="checkbox" checked={seleccionados?.includes(o.clave) ?? false} onChange={() => alternarOpcion(o.clave)} />
                {o.texto.length > 60 ? `${o.texto.slice(0, 60)}…` : o.texto} ({o.cantidad})
              </label>
            ))}
            {seleccionados && (
              <button
                type="button"
                className="text-xs mt-1 font-semibold"
                style={{ color: "var(--acento)" }}
                onClick={() => onCambiarSeleccion(null)}
              >
                Limpiar filtro
              </button>
            )}
          </div>
        </>
      )}
    </>
  );
}

export default function DetalleBdCapex() {
  const [proyectos, setProyectos] = useState<ProyectoCapex[] | null>(null);
  const [archivo, setArchivo] = useState<string>("");
  const [actualizadoEn, setActualizadoEn] = useState<string>("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [busqueda, setBusqueda] = useState("");
  const [gruposSel, setGruposSel] = usePersistedState<string[] | null>("bd-capex-gruposSel", null);
  const [prioridadesSel, setPrioridadesSel] = usePersistedState<string[] | null>("bd-capex-prioridadesSel", null);
  const [avancesSel, setAvancesSel] = usePersistedState<string[] | null>("bd-capex-avancesSel", null);
  const [proyectosSel, setProyectosSel] = usePersistedState<string[] | null>("bd-capex-proyectosSel", null);
  const [detallesSel, setDetallesSel] = usePersistedState<string[] | null>("bd-capex-detallesSel", null);
  const [filtroAbierto, setFiltroAbierto] = useState<Columna | null>(null);
  const [mesesVisibles, setMesesVisibles] = usePersistedState("bd-capex-mesesVisibles", false);
  const [mostrarSoles, setMostrarSoles] = usePersistedState("bd-capex-mostrarSoles", false);
  const [tipoCambio, setTipoCambio] = useTipoCambio();
  const [mostrarFormNuevo, setMostrarFormNuevo] = useState(false);
  const [filaResaltada, setFilaResaltada] = useState<number | null>(null);
  const [orden, setOrden] = useState<{ columna: Columna; asc: boolean }>({
    columna: "proyecto",
    asc: true,
  });
  const [guardandoFila, setGuardandoFila] = useState<number | null>(null);
  const [errorFila, setErrorFila] = useState<{ fila: number; mensaje: string } | null>(null);
  const [mesCierre, setMesCierre] = useMesCierre();
  // Cuántas facturas ya registradas quedan vinculadas a cada fila de proyecto (por su
  // texto exacto de Detalle) — para avisar antes de editar Detalle si eso las va a
  // desvincular (dejan de poder sumarse solas y su Monto se bloquea).
  const [facturasVinculadas, setFacturasVinculadas] = useState<Map<number, number>>(new Map());

  useEffect(() => {
    fetch("/api/facturas", { cache: "no-store" })
      .then((res) => res.json())
      .then((json) => {
        if (!Array.isArray(json.facturas)) return;
        const mapa = new Map<number, number>();
        for (const f of json.facturas as { resolucion: { filaProyecto: number } | null }[]) {
          if (!f.resolucion) continue;
          mapa.set(f.resolucion.filaProyecto, (mapa.get(f.resolucion.filaProyecto) ?? 0) + 1);
        }
        setFacturasVinculadas(mapa);
      })
      .catch(() => {
        // No bloquea la pantalla si esto falla — solo significa que no se podrá avisar.
      });
  }, []);

  /** Aplica un cambio ya guardado con éxito al estado local (sin volver a leer el Excel). */
  function actualizarLocal(fila: number, cambios: Partial<ProyectoCapex>) {
    setProyectos((prev) => prev?.map((x) => (x.filaExcel === fila ? { ...x, ...cambios } : x)) ?? prev);
  }

async function escribirCampoCapex(fila: number, campo: string, valor: number | string) {
    const res = await fetch("/api/capex/celda", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fila, campo, valor }),
    });
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
  }

  async function cambiarAvance(p: ProyectoResuelto, seleccion: string) {
    const avanceAnterior = p.avancePct;
    const statusAnterior = p.status;
    setErrorFila(null);
    setGuardandoFila(p.filaExcel);
    try {
      if (seleccion === "suspendido") {
        // Solo marca Status="Suspendido" — el % de Avance no se toca, así no se pierde
        // el progreso que ya tenía cuando se retome.
        actualizarLocal(p.filaExcel, { status: MARCA_SUSPENDIDO });
        await escribirCampoCapex(p.filaExcel, "status", MARCA_SUSPENDIDO);
      } else {
        const nuevoValor = Number(seleccion);
        actualizarLocal(p.filaExcel, { avancePct: String(nuevoValor), status: "" });
        await escribirCampoCapex(p.filaExcel, "avancePct", nuevoValor);
        // Si venía de "Suspendido", limpia la marca al elegir un % normal de nuevo.
        if (estaSuspendido(statusAnterior)) {
          await escribirCampoCapex(p.filaExcel, "status", "");
        }
      }
    } catch (e) {
      actualizarLocal(p.filaExcel, { avancePct: avanceAnterior, status: statusAnterior });
      setErrorFila({ fila: p.filaExcel, mensaje: (e as Error).message });
    } finally {
      setGuardandoFila(null);
    }
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

  // Si cambia el Proyecto elegido, el Detalle que tenías marcado puede ya no aplicar.
  useEffect(() => {
    setDetallesSel(null);
  }, [proyectosSel]);

  const resueltos = useMemo(() => resolverProyectos(proyectos ?? [], mesCierre), [proyectos, mesCierre]);

  // Un mes ya cerrado (mi < mesCierre) no necesita columna de Proyectado — ya no se usa.
  const mesesConSubcolumnas = useMemo(
    () =>
      NOMBRES_MES.map((nombre, mi) => ({
        nombre,
        mi,
        subcolumnas: mi < mesCierre ? (["real"] as const) : (["real", "proyectado"] as const),
      })),
    [mesCierre]
  );
  const totalSubcolumnasMeses = mesesConSubcolumnas.reduce((a, m) => a + m.subcolumnas.length, 0);

  const grupos = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of resueltos) {
      const clave = p.grupoNegocio || "SIN GRUPO";
      conteo.set(clave, (conteo.get(clave) ?? 0) + 1);
    }
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([valor, cantidad]) => ({ valor, cantidad }));
  }, [resueltos]);
  const prioridades = useMemo(() => prioridadesDisponibles(resueltos), [resueltos]);
  const avances = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of resueltos) {
      const clave = claveAvance(p);
      conteo.set(clave, (conteo.get(clave) ?? 0) + 1);
    }
    return ORDEN_AVANCE.filter((clave) => conteo.has(clave)).map((clave) => ({ clave, cantidad: conteo.get(clave)! }));
  }, [resueltos]);

  const proyectosDisponibles = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of resueltos) conteo.set(p.proyecto, (conteo.get(p.proyecto) ?? 0) + 1);
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([nombre, cantidad]) => ({ nombre, cantidad }));
  }, [resueltos]);

  // Antes de aplicar el filtro de Detalle y la búsqueda de texto — así la lista de
  // Detalle disponible se acota sola según Grupo/Prioridad/Proyecto ya elegidos.
  const preFiltrados = useMemo(() => {
    let filas = resueltos;
    if (gruposSel) filas = filas.filter((p) => gruposSel.includes(p.grupoNegocio || "SIN GRUPO"));
    if (prioridadesSel) filas = filas.filter((p) => prioridadesSel.includes(p.prioridad || "Sin prioridad"));
    if (avancesSel) filas = filas.filter((p) => avancesSel.includes(claveAvance(p)));
    if (proyectosSel) filas = filas.filter((p) => proyectosSel.includes(p.proyecto));
    return filas;
  }, [resueltos, gruposSel, prioridadesSel, avancesSel, proyectosSel]);

  const detallesDisponibles = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of preFiltrados) {
      const d = p.detalle.trim();
      if (!d) continue;
      conteo.set(d, (conteo.get(d) ?? 0) + 1);
    }
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([nombre, cantidad]) => ({ nombre, cantidad }));
  }, [preFiltrados]);

  const filtrados = useMemo(() => {
    let filas = preFiltrados;
    if (detallesSel) filas = filas.filter((p) => detallesSel.includes(p.detalle));
    const q = normalizar(busqueda);
    if (q) {
      filas = filas.filter(
        (p) => normalizar(p.proyecto).includes(q) || normalizar(p.detalle).includes(q) || normalizar(p.responsable).includes(q)
      );
    }
    const copia = [...filas];
    copia.sort((a, b) => {
      // Un proyecto recién creado sin nombre todavía (se completa más adelante) siempre
      // va al final, sin importar la columna o dirección de orden elegida — así no se
      // pierde mezclado al principio de la lista (donde cae por defecto un string vacío
      // al ordenar alfabéticamente).
      const aVacio = !a.proyecto.trim();
      const bVacio = !b.proyecto.trim();
      if (aVacio !== bVacio) return aVacio ? 1 : -1;

      const va = a[orden.columna];
      const vb = b[orden.columna];
      const cmp = typeof va === "number" && typeof vb === "number" ? va - vb : String(va).localeCompare(String(vb));
      return orden.asc ? cmp : -cmp;
    });
    return copia;
  }, [preFiltrados, detallesSel, busqueda, orden]);

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

  function ordenarPor(columna: Columna) {
    setOrden((prev) => (prev.columna === columna ? { columna, asc: !prev.asc } : { columna, asc: false }));
  }

  const hayFiltrosActivos =
    !!busqueda || !!gruposSel || !!prioridadesSel || !!avancesSel || !!proyectosSel || !!detallesSel;

  function limpiarFiltros() {
    setBusqueda("");
    setGruposSel(null);
    setPrioridadesSel(null);
    setAvancesSel(null);
    setProyectosSel(null);
    setDetallesSel(null);
  }

  // Qué columnas tienen filtro en su propio encabezado, y con qué datos/selección/setter —
  // así el <th> genérico de abajo solo necesita mirar aquí para saber si dibuja el ▾.
  const FILTRABLES: Partial<
    Record<
      Columna,
      {
        opciones: OpcionFiltro[];
        seleccion: string[] | null;
        setSeleccion: (v: string[] | null) => void;
        conBusqueda?: boolean;
      }
    >
  > = {
    grupoNegocio: {
      opciones: grupos.map((g) => ({ clave: g.valor, texto: g.valor, cantidad: g.cantidad })),
      seleccion: gruposSel,
      setSeleccion: setGruposSel,
    },
    prioridad: {
      opciones: prioridades.map((p) => ({ clave: p.valor, texto: p.valor, cantidad: p.cantidad })),
      seleccion: prioridadesSel,
      setSeleccion: setPrioridadesSel,
    },
    avancePct: {
      opciones: avances.map((a) => ({ clave: a.clave, texto: ETIQUETAS_AVANCE[a.clave], cantidad: a.cantidad })),
      seleccion: avancesSel,
      setSeleccion: setAvancesSel,
    },
    proyecto: {
      opciones: proyectosDisponibles.map((p) => ({ clave: p.nombre, texto: p.nombre, cantidad: p.cantidad })),
      seleccion: proyectosSel,
      setSeleccion: setProyectosSel,
      conBusqueda: true,
    },
    detalle: {
      opciones: detallesDisponibles.map((d) => ({ clave: d.nombre, texto: d.nombre, cantidad: d.cantidad })),
      seleccion: detallesSel,
      setSeleccion: setDetallesSel,
      conBusqueda: true,
    },
  };

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
        <h2 className="text-lg font-semibold">Detalle BD_CAPEX</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Todas las filas de la hoja BD_CAPEX ({archivo || "…"}), en vivo. {filtrados.length} de{" "}
          {proyectos?.length ?? 0} proyectos.
        </p>
      </div>

      <div className="card p-4 flex flex-col gap-3">
        <div className="flex flex-wrap items-center gap-3">
          <input
            type="text"
            placeholder="Buscar proyecto, detalle o responsable…"
            className="campo"
            style={{ maxWidth: 320 }}
            value={busqueda}
            onChange={(e) => setBusqueda(e.target.value)}
          />
          {hayFiltrosActivos && (
            <button
              type="button"
              className="text-xs font-semibold hover:underline"
              style={{ color: "var(--acento)" }}
              onClick={limpiarFiltros}
              title="Vuelve a mostrar todos los proyectos, sin ningún filtro de por medio"
            >
              Quitar filtros ✕
            </button>
          )}
          <span
            className="chip"
            data-activo={mesesVisibles}
            onClick={() => setMesesVisibles((v) => !v)}
            title="Muestra u oculta el detalle mensual (Ene-Dic) de cada proyecto"
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
          <div className="flex items-center gap-3 ml-auto">
            {actualizadoEn && (
              <span className="text-xs" style={{ color: "var(--texto-suave)" }}>
                Actualizado {new Date(actualizadoEn).toLocaleString("es-PE")}
              </span>
            )}
            <button className="boton-secundario" onClick={() => setMostrarFormNuevo((v) => !v)}>
              {mostrarFormNuevo ? "Cancelar" : "+ Agregar proyecto"}
            </button>
            <button className="boton-primario" onClick={cargar} disabled={cargando}>
              {cargando ? "Actualizando…" : "Actualizar"}
            </button>
          </div>
        </div>

        {mostrarFormNuevo && (
          <FormularioNuevoProyecto
            onCancelar={() => setMostrarFormNuevo(false)}
            onCreado={() => {
              setMostrarFormNuevo(false);
              cargar();
            }}
          />
        )}

        <p className="text-xs" style={{ color: "var(--texto-suave)" }}>
          Filtra por Grupo, Prioridad, Avance, Proyecto o Detalle directamente desde el ▾ en el encabezado de
          cada columna de la tabla.
        </p>
      </div>

      <div className="card p-0 overflow-hidden">
        {/* Scroll propio (no el de toda la página): así el encabezado puede quedar fijo
            arriba con position:sticky mientras se recorre la tabla hacia abajo. El ancho
            de la tabla es siempre el mismo (fijo en px); si no entra en la pantalla,
            aparece scroll horizontal — nunca se recortan los importes. */}
        <div style={{ maxHeight: "75vh", overflow: "auto" }}>
          <table
            className="text-xs border-collapse"
            style={{
              tableLayout: "fixed",
              width: ANCHO_NUMERO + COLUMNAS.reduce((a, c) => a + c.ancho, 0) + (mesesVisibles ? totalSubcolumnasMeses * ANCHO_SUBCOL_MES : 0),
            }}
          >
            <colgroup>
              <col style={{ width: ANCHO_NUMERO }} />
              {COLUMNAS.map((c) => (
                <col key={c.clave} style={{ width: c.ancho }} />
              ))}
              {mesesVisibles &&
                mesesConSubcolumnas.flatMap((m) =>
                  m.subcolumnas.map((s) => <col key={`${m.mi}-${s}`} style={{ width: ANCHO_SUBCOL_MES }} />)
                )}
            </colgroup>
            <thead style={{ position: "sticky", top: 0, zIndex: 10, background: "var(--bg)" }}>
              <tr className="text-left" style={{ color: "var(--texto-suave)", background: "var(--bg)" }}>
                <th rowSpan={mesesVisibles ? 2 : 1} className="py-2 px-3 font-semibold align-top text-center">
                  #
                </th>
                {COLUMNAS.map((c) => {
                  const filtro = FILTRABLES[c.clave];
                  if (!filtro) {
                    return (
                      <th
                        key={c.clave}
                        rowSpan={mesesVisibles ? 2 : 1}
                        className={`py-2 px-3 font-semibold cursor-pointer select-none leading-tight align-top ${c.numerica ? "text-center" : "text-left"}`}
                        style={c.destacada ? ESTILO_DESTACADO : undefined}
                        onClick={() => ordenarPor(c.clave)}
                      >
                        {c.etiqueta}
                        {orden.columna === c.clave ? (orden.asc ? " ▲" : " ▼") : ""}
                      </th>
                    );
                  }
                  const { opciones, seleccion, setSeleccion, conBusqueda } = filtro;
                  return (
                    <th
                      key={c.clave}
                      rowSpan={mesesVisibles ? 2 : 1}
                      className="py-2 px-3 font-semibold select-none leading-tight align-top text-left"
                      style={{ position: "relative" }}
                    >
                      <span className="cursor-pointer" onClick={() => ordenarPor(c.clave)}>
                        {c.etiqueta}
                        {orden.columna === c.clave ? (orden.asc ? " ▲" : " ▼") : ""}
                      </span>
                      <FiltroColumna
                        abierto={filtroAbierto === c.clave}
                        onAbrir={() => setFiltroAbierto(c.clave)}
                        onCerrar={() => setFiltroAbierto(null)}
                        opciones={opciones}
                        seleccionados={seleccion}
                        onCambiarSeleccion={setSeleccion}
                        conBusqueda={conBusqueda}
                      />
                    </th>
                  );
                })}
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
              {filtrados.map((p, i) => {
                const avance = estadoAvance(p.avancePct, p.status);
                const resaltada = filaResaltada === p.filaExcel;
                return (
                  <tr
                    key={i}
                    style={{
                      borderTop: "1px solid var(--borde)",
                      background: resaltada ? COLOR_RESALTADO : undefined,
                      boxShadow: resaltada ? "inset 4px 0 0 var(--acento)" : undefined,
                    }}
                  >
                    <td
                      className="py-1.5 px-3 text-center cursor-pointer font-semibold select-none"
                      style={{ color: resaltada ? "var(--acento-fuerte)" : "var(--texto-suave)" }}
                      title="Clic para resaltar esta fila (útil al exponer en pantalla)"
                      onClick={() => setFilaResaltada(resaltada ? null : p.filaExcel)}
                    >
                      {i + 1}
                    </td>
                    <td className="py-1.5 px-3">
                      <CampoEditable
                        fila={p.filaExcel}
                        campo="proyecto"
                        tipo="texto"
                        valor={p.proyecto}
                        onGuardado={(v) => actualizarLocal(p.filaExcel, { proyecto: String(v) })}
                        confirmarAntes={() => {
                          const n = facturasVinculadas.get(p.filaExcel) ?? 0;
                          if (n === 0) return true;
                          return window.confirm(
                            `Este proyecto tiene ${n} factura${n === 1 ? "" : "s"} registrada${n === 1 ? "" : "s"} que puede${n === 1 ? "" : "n"} depender del nombre del proyecto (cuando el Detalle está vacío, la factura se vincula por el nombre del Proyecto).\n\nSi lo cambias, esa${n === 1 ? "" : "s"} factura${n === 1 ? "" : "s"} podría${n === 1 ? "" : "n"} dejar de vincularse automáticamente a Gasto Real, y su Monto quedaría bloqueado para editarlo más adelante.\n\n¿Deseas continuar de todas formas?`
                          );
                        }}
                      />
                    </td>
                    <td className="py-1.5 px-3" style={{ color: "var(--texto-suave)" }} title={p.detalle}>
                      <CampoEditable
                        fila={p.filaExcel}
                        campo="detalle"
                        tipo="texto"
                        valor={p.detalle}
                        placeholder="—"
                        onGuardado={(v) => actualizarLocal(p.filaExcel, { detalle: String(v) })}
                        confirmarAntes={() => {
                          const n = facturasVinculadas.get(p.filaExcel) ?? 0;
                          if (n === 0) return true;
                          return window.confirm(
                            `Este proyecto tiene ${n} factura${n === 1 ? "" : "s"} registrada${n === 1 ? "" : "s"} vinculada${n === 1 ? "" : "s"} por el texto exacto de este Detalle.\n\nSi lo cambias, esa${n === 1 ? "" : "s"} factura${n === 1 ? "" : "s"} ya no se podrá${n === 1 ? "" : "n"} vincular en automático a Gasto Real, y su Monto quedará bloqueado para editarlo más adelante.\n\n¿Deseas continuar de todas formas?`
                          );
                        }}
                      />
                    </td>
                    <td className="py-1.5 px-3 truncate" title={p.subNegocio}>
                      {p.subNegocio}
                    </td>
                    <td className="py-1.5 px-3 truncate">{p.grupoNegocio}</td>
                    <td className="py-1.5 px-3">
                      <CampoEditable
                        fila={p.filaExcel}
                        campo="prioridad"
                        tipo="texto"
                        valor={p.prioridad}
                        onGuardado={(v) => actualizarLocal(p.filaExcel, { prioridad: String(v) })}
                      />
                    </td>
                    <td className="py-1.5 px-3 overflow-hidden">
                      <select
                        className="rounded-full text-xs font-semibold border-0 cursor-pointer py-0.5 pl-2 pr-6 truncate"
                        style={{
                          background: avance.bg,
                          color: avance.color,
                          opacity: guardandoFila === p.filaExcel ? 0.6 : 1,
                          width: "100%",
                          maxWidth: "100%",
                          boxSizing: "border-box",
                        }}
                        value={estaSuspendido(p.status) ? "suspendido" : (valorAvance(p.avancePct) ?? "")}
                        disabled={guardandoFila === p.filaExcel}
                        onChange={(e) => cambiarAvance(p, e.target.value)}
                      >
                        {OPCIONES_AVANCE.map((o) => (
                          <option key={o.valor} value={o.valor}>
                            {o.texto}
                          </option>
                        ))}
                      </select>
                      {errorFila?.fila === p.filaExcel && (
                        <p className="text-xs mt-1" style={{ color: "var(--peligro)" }}>
                          {errorFila.mensaje}
                        </p>
                      )}
                    </td>
                    <td className="py-1.5 px-3 text-center whitespace-nowrap font-medium" style={ESTILO_DESTACADO}>
                      {moneda2(p.presupuestoAprobado)}
                      {mostrarSoles && (
                        <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                          {soles(p.presupuestoAprobado, tipoCambio)}
                        </span>
                      )}
                    </td>
                    <td className="py-1.5 px-3 text-center whitespace-nowrap font-medium" style={ESTILO_DESTACADO}>
                      {moneda2(p.gastoReal)}
                      {mostrarSoles && (
                        <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                          {soles(p.gastoReal, tipoCambio)}
                        </span>
                      )}
                    </td>
                    <td className="py-1.5 px-3 text-center whitespace-nowrap font-medium" style={ESTILO_DESTACADO}>
                      {moneda2(p.forecast)}
                      {mostrarSoles && (
                        <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                          {soles(p.forecast, tipoCambio)}
                        </span>
                      )}
                    </td>
                    <td
                      className="py-1.5 px-3 text-center whitespace-nowrap font-bold"
                      style={{ ...ESTILO_DESTACADO, color: p.diferencia < 0 ? "var(--peligro)" : "var(--exito)" }}
                    >
                      {moneda2(p.diferencia)}
                    </td>
                    {mesesVisibles &&
                      mesesConSubcolumnas.flatMap((m) =>
                        m.subcolumnas.map((s) => (
                          <td
                            key={`${m.mi}-${s}`}
                            className="py-1.5 px-3"
                            style={{ borderLeft: s === "real" ? "1px solid var(--borde)" : undefined }}
                          >
                            {s === "real" && m.mi < mesCierre ? (
                              // Mes ya cerrado: el Real queda fijo, no se edita más.
                              <span
                                className="block text-center text-xs"
                                style={{ color: "var(--texto-suave)" }}
                                title="Mes cerrado — ya no se edita"
                              >
                                {moneda2(p.real[m.mi])}
                              </span>
                            ) : s === "real" ? (
                              <CampoMontoSumado
                                fila={p.filaExcel}
                                campo={`real:${m.mi}`}
                                valor={p.real[m.mi]}
                                className="text-center text-xs"
                                onGuardado={(nuevo) =>
                                  actualizarLocal(p.filaExcel, {
                                    real: p.real.map((x, j) => (j === m.mi ? nuevo : x)),
                                  })
                                }
                              />
                            ) : (
                              <CampoMontoSumado
                                fila={p.filaExcel}
                                campo={`proyectado:${m.mi}`}
                                valor={p.proyectado[m.mi]}
                                className="text-center text-xs"
                                onGuardado={(nuevo) =>
                                  actualizarLocal(p.filaExcel, {
                                    proyectado: p.proyectado.map((x, j) => (j === m.mi ? nuevo : x)),
                                  })
                                }
                              />
                            )}
                            {mostrarSoles && (
                              <span className="block text-center text-[10px]" style={{ color: "var(--texto-suave)" }}>
                                {soles(s === "real" ? p.real[m.mi] : p.proyectado[m.mi], tipoCambio)}
                              </span>
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
                <td
                  className="py-2 px-3 font-bold"
                  colSpan={CANTIDAD_COLUMNAS_TEXTO}
                  style={{ background: "var(--bg)" }}
                >
                  Total ({filtrados.length} proyectos)
                </td>
                <td className="py-2 px-3 text-center font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.presupuestoAprobado)}
                  {mostrarSoles && (
                    <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.presupuestoAprobado, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.gastoReal)}
                  {mostrarSoles && (
                    <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.gastoReal, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
                  {moneda2(totales.forecast)}
                  {mostrarSoles && (
                    <span className="block text-xs font-normal" style={{ color: "var(--texto-suave)" }}>
                      {soles(totales.forecast, tipoCambio)}
                    </span>
                  )}
                </td>
                <td className="py-2 px-3 text-center font-bold whitespace-nowrap" style={ESTILO_DESTACADO}>
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

const GRUPOS_NEGOCIO = ["EMISIVO", "RECEPTIVO", "TRANSVERSAL"];

/**
 * Formulario para agregar un proyecto nuevo a BD_CAPEX. Solo pide lo mínimo (Grupo,
 * Prioridad, Presupuesto Aprobado — Proyecto, Sub Negocio y Detalle son opcionales, se
 * pueden completar más adelante); el resto (Status, Responsable, Avance, gastos
 * mensuales) se completa después con la edición en línea que ya existe en la tabla.
 */
function FormularioNuevoProyecto({ onCancelar, onCreado }: { onCancelar: () => void; onCreado: () => void }) {
  const [proyecto, setProyecto] = useState("");
  const [subNegocio, setSubNegocio] = useState("");
  const [grupoNegocio, setGrupoNegocio] = useState("");
  const [detalle, setDetalle] = useState("");
  const [prioridad, setPrioridad] = useState("");
  const [presupuestoAprobado, setPresupuestoAprobado] = useState("");
  const [gastoRealInicial, setGastoRealInicial] = useState("");
  const [mesGastoReal, setMesGastoReal] = useState("");
  const [montoProyectado, setMontoProyectado] = useState("");
  const [mesProyectado, setMesProyectado] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const necesitaMes = Number(gastoRealInicial) > 0;
  const necesitaMesProyectado = Number(montoProyectado) > 0;
  const listo =
    grupoNegocio &&
    prioridad.trim() &&
    presupuestoAprobado.trim() &&
    (!necesitaMes || mesGastoReal) &&
    (!necesitaMesProyectado || mesProyectado);

  async function crear() {
    setError(null);
    setEnviando(true);
    try {
      const res = await fetch("/api/capex/agregar-proyecto", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          proyecto: proyecto.trim(),
          subNegocio: subNegocio.trim(),
          grupoNegocio,
          detalle: detalle.trim(),
          prioridad: prioridad.trim(),
          presupuestoAprobado: Number(presupuestoAprobado),
          gastoRealInicial: gastoRealInicial ? Number(gastoRealInicial) : 0,
          mesGastoReal: mesGastoReal ? Number(mesGastoReal) : null,
          montoProyectado: montoProyectado ? Number(montoProyectado) : 0,
          mesProyectado: mesProyectado ? Number(mesProyectado) : null,
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
      <h2 className="font-semibold mb-3 text-sm">Nuevo proyecto</h2>
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div>
          <span className="etiqueta mb-0">Proyecto</span>
          <input
            className="campo"
            value={proyecto}
            onChange={(e) => setProyecto(e.target.value)}
            disabled={enviando}
            placeholder="Se puede completar más adelante"
          />
        </div>
        <div>
          <span className="etiqueta mb-0">Sub Negocio</span>
          <input
            className="campo"
            value={subNegocio}
            onChange={(e) => setSubNegocio(e.target.value)}
            disabled={enviando}
            placeholder="Se puede completar más adelante"
          />
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
        <div>
          <span className="etiqueta mb-0">Gasto Real inicial (USD)</span>
          <input
            type="number"
            className="campo"
            placeholder="0.00"
            value={gastoRealInicial}
            onChange={(e) => setGastoRealInicial(e.target.value)}
            disabled={enviando}
            title="Si el proyecto ya arrancó con algún gasto ejecutado, cárgalo aquí — si no, déjalo vacío."
          />
        </div>
        <div>
          <span className="etiqueta mb-0">Mes del Gasto Real{necesitaMes ? " *" : ""}</span>
          <select
            className="campo"
            value={mesGastoReal}
            onChange={(e) => setMesGastoReal(e.target.value)}
            disabled={enviando || !necesitaMes}
          >
            <option value="">{necesitaMes ? "Elige un mes…" : "—"}</option>
            {NOMBRES_MES_CIERRE.map((nombre, i) => (
              <option key={nombre} value={i + 1}>
                {nombre}
              </option>
            ))}
          </select>
        </div>
        <div>
          <span className="etiqueta mb-0">Monto Proyectado (USD)</span>
          <input
            type="number"
            className="campo"
            placeholder="0.00"
            value={montoProyectado}
            onChange={(e) => setMontoProyectado(e.target.value)}
            disabled={enviando}
            title="Si ya se sabe cuánto se espera gastar y en qué mes, cárgalo aquí — si no, déjalo vacío y se completa después."
          />
        </div>
        <div>
          <span className="etiqueta mb-0">Mes del Monto Proyectado{necesitaMesProyectado ? " *" : ""}</span>
          <select
            className="campo"
            value={mesProyectado}
            onChange={(e) => setMesProyectado(e.target.value)}
            disabled={enviando || !necesitaMesProyectado}
          >
            <option value="">{necesitaMesProyectado ? "Elige un mes…" : "—"}</option>
            {NOMBRES_MES_CIERRE.map((nombre, i) => (
              <option key={nombre} value={i + 1}>
                {nombre}
              </option>
            ))}
          </select>
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