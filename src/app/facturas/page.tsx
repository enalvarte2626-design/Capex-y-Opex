"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE, type FacturaCapex } from "@/lib/capex";
import { moneda2 } from "@/lib/format";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/useTipoCambio";
import { useNivelAcceso } from "@/lib/useNivelAcceso";
import { agruparProveedores, claveNormalizada, mapaRucPorProveedor } from "@/lib/proveedores";
import CampoEditable from "@/components/CampoEditable";

interface ProyectoOpcion {
  filaExcel: number;
  proyecto: string;
  detalle: string;
  grupoNegocio: string;
  responsable: string;
  /** "Sub. Negocio" de BD_CAPEX — el código de empresa (ej. "NM", "CT") de esa línea. */
  subNegocio: string;
}

interface FacturaConResolucion extends FacturaCapex {
  resolucion: { filaProyecto: number; mes: number } | null;
}

interface Respuesta {
  proyectos: ProyectoOpcion[];
  facturas: FacturaConResolucion[];
  actualizadoEn: string;
}

const HOY = () => new Date().toISOString().slice(0, 10);

export default function Facturas() {
  const nivelAcceso = useNivelAcceso();
  const puedeEditar = nivelAcceso === "completo";
  const [datos, setDatos] = useState<Respuesta | null>(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [grupoSel, setGrupoSel] = useState("");
  const [proyectoNombreSel, setProyectoNombreSel] = useState("");
  const [detalleTexto, setDetalleTexto] = useState("");
  const [mostrarSugerenciasDetalle, setMostrarSugerenciasDetalle] = useState(false);

  const [form, setForm] = useState({
    filaProyecto: "",
    mes: String(new Date().getMonth() + 1),
    // Por defecto en Soles SIN IGV (así llegan la mayoría de las facturas locales) — se
    // puede cambiar a Dólares para proveedores que ya facturan en USD directo. El monto
    // que de verdad suma al Gasto Real siempre es en USD (ver "moneda" más abajo).
    moneda: "PEN" as "PEN" | "USD",
    monto: "",
    recurso: "",
    proveedor: "",
    responsable: "",
    numeroFactura: "",
    periodoFacturado: HOY(),
    // Solo aplica a proveedores peruanos (el RUC es un identificador tributario de Perú)
    // — se deja vacío sin problema para proveedores extranjeros.
    ruc: "",
    comentarioExtra: "",
  });
  const [guardando, setGuardando] = useState(false);
  const [mensaje, setMensaje] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);
  const [migrando, setMigrando] = useState(false);
  const [mensajeMigracion, setMensajeMigracion] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);

  /** Migración de datos, un solo uso: copia el mes que ya tenían las facturas antiguas
   *  (texto "Periodo X" en Comentarios) a la nueva columna "Mes Real", y limpia el
   *  comentario. No toca el Gasto Real de BD_CAPEX — esos montos ya están cargados. */
  async function migrarMesReal() {
    const confirmado = window.confirm(
      "Esto corrige el dato del Mes en las facturas ya registradas (moviéndolo de Comentarios a su propia columna). No cambia ningún monto del presupuesto. Puede tardar hasta un minuto si hay muchas facturas — no cierres esta pestaña mientras dice \"Corrigiendo…\". ¿Continuar?"
    );
    if (!confirmado) return;
    setMigrando(true);
    setMensajeMigracion(null);
    try {
      const res = await fetch("/api/facturas/migrar-mes-real", { method: "POST" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo migrar.");
      setMensajeMigracion({
        tipo: "ok",
        texto: `Listo: ${json.migradas} de ${json.totalFacturas} facturas corregidas.${
          json.sinMesDetectado?.length ? ` ${json.sinMesDetectado.length} sin mes detectable (corrígelas a mano en la columna Mes).` : ""
        }`,
      });
      await cargar();
    } catch (e) {
      setMensajeMigracion({ tipo: "error", texto: (e as Error).message });
    } finally {
      setMigrando(false);
    }
  }

  /** Aplica un cambio ya guardado con éxito a una factura, sin volver a leer el Excel. */
  function actualizarFacturaLocal(filaExcel: number, cambios: Partial<FacturaConResolucion>) {
    setDatos((prev) =>
      prev
        ? { ...prev, facturas: prev.facturas.map((f) => (f.filaExcel === filaExcel ? { ...f, ...cambios } : f)) }
        : prev
    );
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/facturas", { cache: "no-store" });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo cargar.");
      setDatos(json);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
  }, []);

  // Proveedores que ya aparecen en facturas registradas — para sugerirlos en el
  // desplegable y no tener que volver a tipear el nombre cada mes. `agruparProveedores`
  // junta variaciones de escritura del mismo proveedor: mayúsculas/espacios distintos
  // ("Go daddy" = "Go Daddy") y también razón social con/sin sufijo ("Metrica" =
  // "Metrica Sac") — ver el comentario en lib/proveedores.ts para el criterio exacto.
  const proveedoresConocidos = useMemo(
    () => agruparProveedores((datos?.facturas ?? []).map((f) => f.recurso)),
    [datos]
  );

  // RUC más usado con cada proveedor ya registrado — para autocompletarlo apenas se
  // elige un proveedor conocido. Uno nuevo (o sin RUC en ningún registro anterior) no
  // aparece acá, así que el campo queda vacío para completarlo a mano.
  const rucPorProveedor = useMemo(
    () => mapaRucPorProveedor((datos?.facturas ?? []).map((f) => ({ proveedor: f.recurso, ruc: f.ruc }))),
    [datos]
  );

  // Grupos de Negocio disponibles — elegir uno acota tanto el Proyecto como el Detalle
  // de abajo (opcional: se puede saltar directo a buscar por Proyecto o Detalle).
  const grupos = useMemo(() => {
    const set = new Set((datos?.proyectos ?? []).map((p) => p.grupoNegocio).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "es"));
  }, [datos]);

  const proyectosDelGrupo = useMemo(
    () => (datos?.proyectos ?? []).filter((p) => !grupoSel || p.grupoNegocio === grupoSel),
    [datos, grupoSel]
  );

  const nombresProyecto = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of proyectosDelGrupo) conteo.set(p.proyecto, (conteo.get(p.proyecto) ?? 0) + 1);
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([nombre, cantidad]) => ({ nombre, cantidad }));
  }, [proyectosDelGrupo]);

  // El Detalle se puede buscar directo (sin elegir Proyecto antes): se acota al Grupo
  // elegido (si hay) y al Proyecto (si ya se eligió uno), pero siempre se puede escribir
  // parte de un Detalle O de un nombre de Proyecto para encontrarlo — al elegir una
  // sugerencia, autocompleta Proyecto (y Grupo) solo.
  const opcionesDetalle = useMemo(() => {
    const lista = proyectosDelGrupo.filter((p) => !proyectoNombreSel || p.proyecto === proyectoNombreSel);
    const vistos = new Set<string>();
    const resultado: { etiqueta: string; buscable: string; filaExcel: number }[] = [];
    for (const p of lista) {
      const detalle = p.detalle.trim() || "(sin detalle)";
      const clave = `${p.proyecto}|||${detalle}`;
      if (vistos.has(clave)) continue;
      vistos.add(clave);
      resultado.push({
        // Sin Proyecto elegido todavía, el nombre del Proyecto va junto en lo que se
        // muestra, para no confundir Detalles iguales de proyectos distintos.
        etiqueta: proyectoNombreSel ? detalle : `${detalle} — ${p.proyecto}`,
        // Para filtrar mientras se escribe: siempre incluye Detalle y Proyecto, aunque
        // no se muestren juntos, así buscar por cualquiera de los dos encuentra la fila.
        buscable: `${detalle} ${p.proyecto}`.toLowerCase(),
        filaExcel: p.filaExcel,
      });
    }
    return resultado.sort((a, b) => a.etiqueta.localeCompare(b.etiqueta, "es"));
  }, [proyectosDelGrupo, proyectoNombreSel]);

  const sugerenciasDetalle = useMemo(() => {
    const texto = detalleTexto.trim().toLowerCase();
    const lista = texto ? opcionesDetalle.filter((o) => o.buscable.includes(texto)) : opcionesDetalle;
    return lista.slice(0, 30);
  }, [opcionesDetalle, detalleTexto]);

  function elegirGrupo(valor: string) {
    setGrupoSel(valor);
    setProyectoNombreSel("");
    setDetalleTexto("");
    setForm((prev) => ({ ...prev, filaProyecto: "" }));
  }

  function elegirProyecto(valor: string) {
    setProyectoNombreSel(valor);
    setDetalleTexto("");
    setForm((prev) => ({ ...prev, filaProyecto: "" }));
  }

  // Al hacer clic en una sugerencia, identifica de una sola vez la fila de BD_CAPEX y
  // autocompleta Proyecto y Grupo — así se puede buscar por Detalle o por Proyecto sin
  // tener que elegir nada antes.
  function elegirSugerenciaDetalle(opcion: { etiqueta: string; filaExcel: number }) {
    const p = datos?.proyectos.find((x) => x.filaExcel === opcion.filaExcel);
    if (!p) return;
    setDetalleTexto(opcion.etiqueta);
    setMostrarSugerenciasDetalle(false);
    setGrupoSel(p.grupoNegocio);
    setProyectoNombreSel(p.proyecto);
    setForm((prev) => ({
      ...prev,
      filaProyecto: String(p.filaExcel),
      // Siempre el Responsable y la Empresa de la fila elegida (no solo si estaban
      // vacíos) — así quedan por defecto según el proyecto/detalle cada vez que cambia
      // la selección, y se pueden seguir corrigiendo a mano después si hace falta.
      responsable: p.responsable || prev.responsable,
      proveedor: p.subNegocio || prev.proveedor,
    }));
  }

  // Escribir invalida la selección anterior (si había) — hay que volver a elegir de la
  // lista de sugerencias para que quede identificada la fila de BD_CAPEX.
  function cambiarTextoDetalle(texto: string) {
    setDetalleTexto(texto);
    setMostrarSugerenciasDetalle(true);
    setForm((prev) => (prev.filaProyecto ? { ...prev, filaProyecto: "" } : prev));
  }

  const proyectoElegido = datos?.proyectos.find((p) => String(p.filaExcel) === form.filaProyecto);

  // Solo para mostrar el equivalente en pantalla mientras se escribe — el backend hace
  // su propio cálculo con el mismo tipo de cambio, así que esto es únicamente una vista
  // previa, nunca lo que de verdad se guarda.
  const montoNum = Number(form.monto);
  const hayMontoValido = Number.isFinite(montoNum) && montoNum !== 0;
  const montoUsdPrevio =
    form.moneda === "PEN" && hayMontoValido ? Math.round((montoNum / TIPO_CAMBIO_POR_DEFECTO) * 100) / 100 : null;
  const montoSolesPrevio =
    form.moneda === "USD" && hayMontoValido ? Math.round(montoNum * TIPO_CAMBIO_POR_DEFECTO * 100) / 100 : null;

  // Al elegir (o terminar de escribir) un Proveedor ya usado antes, autocompleta el RUC
  // con el que más veces se registró para ese mismo proveedor. Un proveedor nuevo (o que
  // nunca se registró con RUC) simplemente no tiene coincidencia — el campo queda vacío
  // para completarlo a mano, tal como pide el flujo.
  function cambiarProveedor(valor: string) {
    const rucSugerido = rucPorProveedor.get(claveNormalizada(valor));
    setForm((prev) => ({
      ...prev,
      recurso: valor,
      ruc: rucSugerido ?? prev.ruc,
    }));
  }

  function actualizarCampo(campo: keyof typeof form, valor: string) {
    setForm((prev) => ({ ...prev, [campo]: valor }));
  }

  async function registrar(e: React.FormEvent) {
    e.preventDefault();
    if (!proyectoElegido) {
      setMensaje({ tipo: "error", texto: "Elige un proyecto." });
      return;
    }
    const monto = Number(form.monto);
    // Negativo se permite a propósito: es como se registra un descuento o nota de
    // crédito (resta del Gasto Real en vez de sumar). Solo 0 no tiene sentido.
    if (!Number.isFinite(monto) || monto === 0) {
      setMensaje({
        tipo: "error",
        texto:
          form.moneda === "PEN"
            ? "El monto en Soles no puede ser 0 (usa negativo para un descuento o nota de crédito)."
            : "El monto en dólares no puede ser 0 (usa negativo para un descuento o nota de crédito).",
      });
      return;
    }

    const mesTexto = NOMBRES_MES_CIERRE[Number(form.mes) - 1];
    const montoUsd = form.moneda === "USD" ? monto : Math.round((monto / TIPO_CAMBIO_POR_DEFECTO) * 100) / 100;
    const esDescuento = monto < 0;
    const descripcionMonto =
      form.moneda === "PEN"
        ? `S/ ${monto.toFixed(2)} (sin IGV) — equivale a ${moneda2(montoUsd)} al tipo de cambio ${TIPO_CAMBIO_POR_DEFECTO}`
        : `${moneda2(monto)}`;
    const confirmado = window.confirm(
      `¿Registrar ${esDescuento ? "un descuento/nota de crédito" : "factura"} de ${descripcionMonto} para "${proyectoElegido.proyecto} — ${proyectoElegido.detalle || "(sin detalle)"}", período ${mesTexto}? Esto ${esDescuento ? "resta" : "suma"} ${moneda2(Math.abs(montoUsd))} al Gasto Real de ${mesTexto} en BD_CAPEX y agrega una fila en la hoja de facturas.`
    );
    if (!confirmado) return;

    setGuardando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/facturas/registrar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          filaProyecto: Number(form.filaProyecto),
          mes: Number(form.mes),
          moneda: form.moneda,
          monto,
          recurso: form.recurso,
          proveedor: form.proveedor,
          responsable: form.responsable,
          numeroFactura: form.numeroFactura,
          periodoFacturado: form.periodoFacturado,
          ruc: form.ruc || undefined,
          comentarioExtra: form.comentarioExtra || undefined,
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo registrar la factura.");
      setMensaje({
        tipo: "ok",
        texto: `Factura registrada (${moneda2(json.monto)} al tipo de cambio ${json.tipoCambio}). Gasto Real de ${mesTexto}: ${moneda2(json.gastoRealAnterior)} → ${moneda2(json.gastoRealNuevo)}.`,
      });
      setForm((prev) => ({
        ...prev,
        monto: "",
        numeroFactura: "",
        ruc: "",
        comentarioExtra: "",
      }));
      await cargar();
    } catch (e) {
      setMensaje({ tipo: "error", texto: (e as Error).message });
    } finally {
      setGuardando(false);
    }
  }

  if (cargando && !datos) {
    return <p style={{ color: "var(--texto-suave)" }}>Cargando datos desde SharePoint…</p>;
  }

  if (error) {
    return (
      <div className="card p-6">
        <p className="font-semibold mb-1" style={{ color: "var(--peligro)" }}>
          No se pudo cargar
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
        <h2 className="text-lg font-semibold">Facturas</h2>
      </div>

      {!puedeEditar && (
        <div className="card p-4 text-sm" style={{ color: "var(--texto-suave)" }}>
          Estás viendo Facturas en modo de solo lectura — puedes ver la tabla, pero no registrar ni editar facturas.
        </div>
      )}

      {puedeEditar && (
      <form onSubmit={registrar} className="card p-4 flex flex-col gap-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="etiqueta">Grupo de Negocio (opcional, acota Proyecto/Detalle)</label>
            <select className="campo" value={grupoSel} onChange={(e) => elegirGrupo(e.target.value)}>
              <option value="">Todos</option>
              {grupos.map((g) => (
                <option key={g} value={g}>
                  {g}
                </option>
              ))}
            </select>
          </div>
          <div />
          <div>
            <label className="etiqueta">Proyecto</label>
            <select className="campo" value={proyectoNombreSel} onChange={(e) => elegirProyecto(e.target.value)} required>
              <option value="">Selecciona un proyecto…</option>
              {nombresProyecto.map((p) => (
                <option key={p.nombre} value={p.nombre}>
                  {p.nombre} ({p.cantidad})
                </option>
              ))}
            </select>
          </div>
          <div style={{ position: "relative" }}>
            <label className="etiqueta">Detalle (también se puede buscar por nombre de Proyecto)</label>
            <input
              type="text"
              className="campo"
              value={detalleTexto}
              onChange={(e) => cambiarTextoDetalle(e.target.value)}
              onFocus={() => setMostrarSugerenciasDetalle(true)}
              onBlur={() => setTimeout(() => setMostrarSugerenciasDetalle(false), 150)}
              placeholder="Escribe para buscar…"
              autoComplete="off"
              required
            />
            {mostrarSugerenciasDetalle && sugerenciasDetalle.length > 0 && (
              <ul
                className="card"
                style={{
                  position: "absolute",
                  zIndex: 20,
                  top: "100%",
                  left: 0,
                  right: 0,
                  marginTop: 2,
                  maxHeight: 220,
                  overflowY: "auto",
                  padding: "0.25rem 0",
                }}
              >
                {sugerenciasDetalle.map((o) => (
                  <li key={o.filaExcel}>
                    <button
                      type="button"
                      className="text-xs w-full text-left"
                      style={{ padding: "0.35rem 0.6rem" }}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => elegirSugerenciaDetalle(o)}
                    >
                      {o.etiqueta}
                    </button>
                  </li>
                ))}
              </ul>
            )}
            {proyectoElegido ? (
              <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
                Proyecto: <strong>{proyectoElegido.proyecto}</strong> ({proyectoElegido.grupoNegocio})
              </p>
            ) : (
              <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
                Elige una sugerencia de la lista para identificar la fila.
              </p>
            )}
          </div>

          <div>
            <label className="etiqueta">Mes Real (mes al que pertenece el gasto — carga aquí al presupuesto)</label>
            <select className="campo" value={form.mes} onChange={(e) => actualizarCampo("mes", e.target.value)} required>
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  {nombre}
                </option>
              ))}
            </select>
            <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
              Puede ser distinto a la fecha de emisión del comprobante — ej. una factura
              emitida en agosto por un servicio de julio va aquí en Julio.
            </p>
          </div>

          <div>
            <label className="etiqueta">Moneda</label>
            <select
              className="campo"
              value={form.moneda}
              onChange={(e) => actualizarCampo("moneda", e.target.value)}
            >
              <option value="PEN">Soles (sin IGV)</option>
              <option value="USD">Dólares (USD)</option>
            </select>
          </div>
          <div>
            <label className="etiqueta">{form.moneda === "PEN" ? "Monto en Soles (sin IGV)" : "Monto en Dólares (USD)"}</label>
            <input
              type="number"
              step="0.01"
              className="campo"
              value={form.monto}
              onChange={(e) => actualizarCampo("monto", e.target.value)}
              placeholder="Negativo = descuento o nota de crédito"
              required
            />
            {montoUsdPrevio != null && (
              <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
                ≈ {moneda2(montoUsdPrevio)} al tipo de cambio {TIPO_CAMBIO_POR_DEFECTO}
              </p>
            )}
            {montoSolesPrevio != null && (
              <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
                ≈ S/ {montoSolesPrevio.toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} al tipo de
                cambio {TIPO_CAMBIO_POR_DEFECTO}
              </p>
            )}
          </div>
          <div>
            <label className="etiqueta">N° Factura</label>
            <input
              type="text"
              className="campo"
              value={form.numeroFactura}
              onChange={(e) => actualizarCampo("numeroFactura", e.target.value)}
            />
          </div>
          <div>
            <label className="etiqueta">RUC (opcional, solo proveedores peruanos)</label>
            <input type="text" className="campo" value={form.ruc} onChange={(e) => actualizarCampo("ruc", e.target.value)} />
          </div>

          <div>
            <label className="etiqueta">Proveedor</label>
            <input
              type="text"
              list="proveedores-capex"
              className="campo"
              value={form.recurso}
              onChange={(e) => cambiarProveedor(e.target.value)}
              placeholder="Elige uno ya usado o escribe uno nuevo"
            />
            <datalist id="proveedores-capex">
              {proveedoresConocidos.map((p) => (
                <option key={p} value={p} />
              ))}
            </datalist>
          </div>
          <div>
            <label className="etiqueta">Empresa (código)</label>
            <input type="text" className="campo" value={form.proveedor} onChange={(e) => actualizarCampo("proveedor", e.target.value)} />
          </div>

          <div>
            <label className="etiqueta">Responsable</label>
            <input
              type="text"
              className="campo"
              value={form.responsable}
              onChange={(e) => actualizarCampo("responsable", e.target.value)}
            />
          </div>
          <div>
            <label className="etiqueta">Fecha de emisión del comprobante</label>
            <input
              type="date"
              className="campo"
              value={form.periodoFacturado}
              onChange={(e) => actualizarCampo("periodoFacturado", e.target.value)}
              required
            />
          </div>

          <div className="sm:col-span-2">
            <label className="etiqueta">Comentario adicional (opcional)</label>
            <input
              type="text"
              className="campo"
              placeholder='Se guarda junto a "Periodo {mes}" en la columna Comentarios'
              value={form.comentarioExtra}
              onChange={(e) => actualizarCampo("comentarioExtra", e.target.value)}
            />
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button type="submit" className="boton-primario" disabled={guardando}>
            {guardando ? "Registrando…" : "Registrar factura"}
          </button>
          {mensaje && (
            <span className="text-sm" style={{ color: mensaje.tipo === "error" ? "var(--peligro)" : "var(--exito)" }}>
              {mensaje.texto}
            </span>
          )}
        </div>
      </form>
      )}

      <div className="card p-0 overflow-hidden">
        <div className="p-4 pb-0 flex items-center justify-between gap-3 flex-wrap">
          <h3 className="font-semibold">Últimas facturas registradas</h3>
          <div className="flex items-center gap-3 flex-wrap">
            {puedeEditar && (
              <button type="button" className="boton-secundario text-xs" onClick={migrarMesReal} disabled={migrando}>
                {migrando ? "Corrigiendo…" : "Corregir Mes de facturas antiguas (una vez)"}
              </button>
            )}
            <a href="/api/facturas/exportar" className="boton-secundario" download>
              Descargar reporte (Excel)
            </a>
          </div>
        </div>
        {mensajeMigracion && (
          <p
            className="px-4 pt-2 text-xs"
            style={{ color: mensajeMigracion.tipo === "error" ? "var(--peligro)" : "var(--exito)" }}
          >
            {mensajeMigracion.texto}
          </p>
        )}
        <p className="px-4 pt-1 text-xs" style={{ color: "var(--texto-suave)" }}>
          El Excel descargado viene ordenado por Proyecto.
          {puedeEditar &&
            " Todos los campos de la tabla de abajo son editables directo — se guardan al salir del campo. Cambiar el Mes solo corrige la etiqueta de esta factura, no vuelve a sumar ni restar del presupuesto. El Monto solo se puede corregir cuando la app identifica con certeza a qué fila/mes de BD_CAPEX corresponde (si no, sale de solo lectura, con una nota)."}
        </p>
        <div className="overflow-x-auto p-4">
          <table className="border-collapse" style={{ tableLayout: "fixed", width: "100%", minWidth: 1400 }}>
            <colgroup>
              <col style={{ width: 110 }} />
              <col style={{ width: 90 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 100 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 240 }} />
              <col style={{ width: 110 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 110 }} />
              <col style={{ width: 220 }} />
            </colgroup>
            <thead>
              <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                <th className="py-2 pr-3 text-xs font-semibold" title="Fecha de emisión del comprobante — no necesariamente el mes al que se cargó el gasto en el presupuesto">Fecha emisión</th>
                <th className="py-2 pr-3 text-xs font-semibold" title="Mes al que se cargó el gasto en el presupuesto — puede ser distinto a la Fecha emisión">Mes</th>
                <th className="py-2 pr-3 text-xs font-semibold">Proveedor</th>
                <th className="py-2 pr-3 text-xs font-semibold">Empresa</th>
                <th className="py-2 pr-3 text-xs font-semibold">Responsable</th>
                <th className="py-2 pr-3 text-xs font-semibold">Proyecto</th>
                <th className="py-2 pr-3 text-xs font-semibold text-right">Monto (USD)</th>
                <th className="py-2 pr-3 text-xs font-semibold text-right">Soles (sin IGV)</th>
                <th className="py-2 pr-3 text-xs font-semibold">N° Factura</th>
                <th className="py-2 pr-3 text-xs font-semibold">RUC</th>
                <th className="py-2 pr-3 text-xs font-semibold">Comentarios</th>
              </tr>
            </thead>
            <tbody>
              {(datos?.facturas ?? []).slice(0, 25).map((f) => (
                <tr key={f.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td className="py-1.5 pr-3">
                    <CampoFecha
                      factura={f}
                      onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)}
                      soloLectura={!puedeEditar}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoMes factura={f} onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)} soloLectura={!puedeEditar} />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="recurso"
                      tipo="texto"
                      valor={f.recurso}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { recurso: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="proveedor"
                      tipo="texto"
                      valor={f.proveedor}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { proveedor: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="responsable"
                      tipo="texto"
                      valor={f.responsable}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { responsable: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3 truncate" style={{ color: "var(--texto-suave)" }} title={f.proyecto}>
                    {f.proyecto}
                  </td>
                  <td className="py-1.5 pr-3">
                    <MontoFactura
                      factura={f}
                      onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)}
                      soloLectura={!puedeEditar}
                    />
                  </td>
                  <td
                    className="py-1.5 pr-3 text-right text-xs"
                    style={{ color: "var(--texto-suave)" }}
                    title={f.montoSolesEsCalculado ? "Factura ingresada en Dólares — equivalente en Soles solo de referencia" : undefined}
                  >
                    {f.montoSoles != null
                      ? `${f.montoSolesEsCalculado ? "≈ " : ""}S/ ${f.montoSoles.toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}${f.tipoCambio ? ` (TC ${f.tipoCambio})` : ""}`
                      : "—"}
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="numeroFactura"
                      tipo="texto"
                      valor={f.numeroFactura}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { numeroFactura: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="ruc"
                      tipo="texto"
                      valor={f.ruc}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { ruc: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="comentarios"
                      tipo="texto"
                      valor={f.comentarios}
                      endpoint="/api/facturas/editar-campo"
                      soloLectura={!puedeEditar}
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { comentarios: String(v) })}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/** Edita el Periodo facturado con un selector de fecha real (no texto libre). */
function CampoFecha({
  factura,
  onGuardado,
  soloLectura,
}: {
  factura: FacturaConResolucion;
  onGuardado: (cambios: Partial<FacturaConResolucion>) => void;
  soloLectura?: boolean;
}) {
  const [valor, setValor] = useState(factura.periodoFacturadoISO);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => setValor(factura.periodoFacturadoISO), [factura.periodoFacturadoISO]);

  if (soloLectura) {
    return <span className="text-xs">{factura.periodoFacturado || "—"}</span>;
  }

  async function guardar() {
    if (!valor || valor === factura.periodoFacturadoISO) return;
    setError(null);
    setGuardando(true);
    try {
      const res = await fetch("/api/facturas/editar-campo", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fila: factura.filaExcel, campo: "periodoFacturado", valor }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
      const fecha = new Date(valor);
      onGuardado({
        periodoFacturadoISO: valor,
        periodoFacturado: fecha.toLocaleDateString("es-PE", { year: "numeric", month: "2-digit", day: "2-digit" }),
      });
    } catch (e) {
      setValor(factura.periodoFacturadoISO);
      setError((e as Error).message);
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div>
      <input
        type="date"
        className="text-xs"
        style={{ background: "transparent", border: "1px solid transparent", width: "100%", opacity: guardando ? 0.6 : 1 }}
        value={valor}
        disabled={guardando}
        onChange={(e) => setValor(e.target.value)}
        onBlur={guardar}
      />
      {error && (
        <p className="text-xs mt-0.5" style={{ color: "var(--peligro)" }}>
          {error}
        </p>
      )}
    </div>
  );
}

/**
 * Corrige el Mes al que pertenece el gasto (independiente de la Fecha de emisión) —
 * solo cambia esta etiqueta/dato; NUNCA vuelve a sumar ni restar del Gasto Real de
 * BD_CAPEX (ese ajuste, si hiciera falta, se hace aparte con el campo Monto).
 */
function CampoMes({
  factura,
  onGuardado,
  soloLectura,
}: {
  factura: FacturaConResolucion;
  onGuardado: (cambios: Partial<FacturaConResolucion>) => void;
  soloLectura?: boolean;
}) {
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  if (soloLectura) {
    return <span className="text-xs">{factura.mesReal ? NOMBRES_MES_CIERRE[factura.mesReal - 1] : "—"}</span>;
  }

  async function guardar(valor: string) {
    const mes = Number(valor);
    if (!Number.isInteger(mes) || mes === factura.mesReal) return;
    setError(null);
    setGuardando(true);
    try {
      const res = await fetch("/api/facturas/editar-campo", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fila: factura.filaExcel, campo: "mesReal", valor: mes }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
      onGuardado({ mesReal: mes });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div>
      <select
        className="text-xs"
        style={{ background: "transparent", border: "1px solid transparent", width: "100%", opacity: guardando ? 0.6 : 1 }}
        value={factura.mesReal ?? ""}
        disabled={guardando}
        onChange={(e) => guardar(e.target.value)}
      >
        <option value="" disabled>
          —
        </option>
        {NOMBRES_MES_CIERRE.map((nombre, i) => (
          <option key={nombre} value={i + 1}>
            {nombre}
          </option>
        ))}
      </select>
      {error && (
        <p className="text-xs mt-0.5" style={{ color: "var(--peligro)" }}>
          {error}
        </p>
      )}
    </div>
  );
}

/**
 * Edita el Monto de una factura ya registrada. Solo permite guardar cuando la app puede
 * identificar con certeza a qué fila/mes de BD_CAPEX corresponde (mismo texto que se
 * escribió al registrarla) — si no, queda de solo lectura, porque corregir el monto sin
 * eso significaría no poder ajustar el Gasto Real, o arriesgar a tocar la celda
 * equivocada.
 */
function MontoFactura({
  factura,
  onGuardado,
  soloLectura,
}: {
  factura: FacturaConResolucion;
  onGuardado: (cambios: Partial<FacturaConResolucion>) => void;
  soloLectura?: boolean;
}) {
  const [valorLocal, setValorLocal] = useState(String(factura.monto));
  const [enfocado, setEnfocado] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!enfocado) setValorLocal(String(factura.monto));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [factura.monto]);

  if (soloLectura) {
    return <span className="text-right block text-xs">{moneda2(factura.monto)}</span>;
  }

  if (!factura.resolucion) {
    return (
      <span
        className="text-right block text-xs"
        style={{ color: "var(--texto-suave)" }}
        title="No se puede identificar con certeza a qué proyecto/mes de BD_CAPEX corresponde esta factura (formato antiguo o distinto al que usa este módulo), así que el monto no se puede corregir aquí."
      >
        {moneda2(factura.monto)} 🔒
      </span>
    );
  }

  async function guardar() {
    setEnfocado(false);
    const nuevo = Number(valorLocal);
    if (!Number.isFinite(nuevo) || nuevo === 0 || nuevo === factura.monto) {
      setValorLocal(String(factura.monto));
      return;
    }
    setError(null);
    setGuardando(true);
    try {
      const res = await fetch("/api/facturas/editar-monto", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filaFactura: factura.filaExcel, montoNuevo: nuevo }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
      onGuardado({ monto: nuevo });
    } catch (e) {
      setValorLocal(String(factura.monto));
      setError((e as Error).message);
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div>
      <input
        type="text"
        inputMode="decimal"
        className="text-right text-xs"
        style={{ background: "transparent", border: "1px solid transparent", width: "100%", opacity: guardando ? 0.6 : 1 }}
        value={enfocado ? valorLocal : moneda2(Number(valorLocal) || 0)}
        disabled={guardando}
        onChange={(e) => setValorLocal(e.target.value)}
        onBlur={guardar}
        onKeyDown={(e) => {
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        }}
        onFocus={(e) => {
          setEnfocado(true);
          e.target.style.border = "1px solid var(--acento)";
        }}
      />
      {error && (
        <p className="text-xs mt-0.5" style={{ color: "var(--peligro)" }}>
          {error}
        </p>
      )}
    </div>
  );
}