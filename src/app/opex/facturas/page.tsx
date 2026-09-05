"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE } from "@/lib/capex";
import { moneda2 } from "@/lib/format";
import { TIPO_CAMBIO_POR_DEFECTO } from "@/lib/useTipoCambio";
import CampoEditable from "@/components/CampoEditable";
import type { FacturaOpex } from "@/lib/opex-parse";

interface LineaOpcion {
  filaExcel: number;
  grupoGasto: string;
  subgrupoGasto: string;
  lineaGasto: string;
  responsable: string;
}

interface Respuesta {
  lineas: LineaOpcion[];
  facturas: FacturaOpex[];
  actualizadoEn: string;
}

export default function FacturasOpex() {
  const [datos, setDatos] = useState<Respuesta | null>(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [grupoSel, setGrupoSel] = useState("");
  const [subgrupoSel, setSubgrupoSel] = useState("");
  const [lineaSel, setLineaSel] = useState("");

  const [form, setForm] = useState({
    mes: String(new Date().getMonth() + 1),
    // Se ingresa en Soles SIN IGV — el equivalente en USD que de verdad mueve el
    // presupuesto se calcula solo, con el tipo de cambio fijo de la app, y nunca se
    // pide directo: así toda factura queda convertida con el mismo criterio.
    montoSoles: "",
    proveedor: "",
    numeroComprobante: "",
    comentario: "",
  });
  const [guardando, setGuardando] = useState(false);
  const [mensaje, setMensaje] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);

  // Solo para mostrar el equivalente en pantalla mientras se escribe — el backend hace
  // su propio cálculo con el mismo tipo de cambio, así que esto es únicamente una
  // vista previa, nunca lo que de verdad se guarda.
  const montoSolesNum = Number(form.montoSoles);
  const montoUsdPrevio =
    Number.isFinite(montoSolesNum) && montoSolesNum > 0
      ? Math.round((montoSolesNum / TIPO_CAMBIO_POR_DEFECTO) * 100) / 100
      : null;

  function actualizarFacturaLocal(filaExcel: number, cambios: Partial<FacturaOpex>) {
    setDatos((prev) =>
      prev ? { ...prev, facturas: prev.facturas.map((f) => (f.filaExcel === filaExcel ? { ...f, ...cambios } : f)) } : prev
    );
  }

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      const res = await fetch("/api/opex/facturas", { cache: "no-store" });
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

  const grupos = useMemo(() => Array.from(new Set((datos?.lineas ?? []).map((l) => l.grupoGasto))).sort(), [datos]);

  const subgrupos = useMemo(() => {
    const lista = (datos?.lineas ?? []).filter((l) => !grupoSel || l.grupoGasto === grupoSel);
    return Array.from(new Set(lista.map((l) => l.subgrupoGasto))).sort();
  }, [datos, grupoSel]);

  const lineasDisponibles = useMemo(() => {
    return (datos?.lineas ?? []).filter(
      (l) => (!grupoSel || l.grupoGasto === grupoSel) && (!subgrupoSel || l.subgrupoGasto === subgrupoSel)
    );
  }, [datos, grupoSel, subgrupoSel]);

  // Al cambiar Grupo, el Subgrupo/Línea elegidos antes pueden ya no aplicar.
  useEffect(() => {
    setSubgrupoSel("");
    setLineaSel("");
  }, [grupoSel]);
  useEffect(() => {
    setLineaSel("");
  }, [subgrupoSel]);

  const lineaElegida = datos?.lineas.find((l) => String(l.filaExcel) === lineaSel);

  async function registrar(e: React.FormEvent) {
    e.preventDefault();
    if (!lineaElegida) {
      setMensaje({ tipo: "error", texto: "Elige una línea de gasto." });
      return;
    }
    const montoSoles = Number(form.montoSoles);
    if (!Number.isFinite(montoSoles) || montoSoles <= 0) {
      setMensaje({ tipo: "error", texto: "El monto en Soles debe ser mayor a 0." });
      return;
    }
    if (!form.proveedor.trim()) {
      setMensaje({ tipo: "error", texto: "Falta el Proveedor." });
      return;
    }

    const mesTexto = NOMBRES_MES_CIERRE[Number(form.mes) - 1];
    const montoUsd = Math.round((montoSoles / TIPO_CAMBIO_POR_DEFECTO) * 100) / 100;
    const confirmado = window.confirm(
      `¿Registrar factura de S/ ${montoSoles.toFixed(2)} (sin IGV) — equivale a ${moneda2(montoUsd)} al tipo de cambio ${TIPO_CAMBIO_POR_DEFECTO} — para "${lineaElegida.lineaGasto}", período ${mesTexto}? Esto suma ${moneda2(montoUsd)} al Gasto Real de ${mesTexto} en Presupuesto 2026.`
    );
    if (!confirmado) return;

    setGuardando(true);
    setMensaje(null);
    try {
      const res = await fetch("/api/opex/facturas/registrar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          filaPresupuesto: lineaElegida.filaExcel,
          mes: Number(form.mes),
          montoSoles,
          proveedor: form.proveedor,
          numeroComprobante: form.numeroComprobante,
          comentario: form.comentario || undefined,
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo registrar la factura.");
      setMensaje({
        tipo: "ok",
        texto: `Factura registrada (${moneda2(json.monto)} al tipo de cambio ${json.tipoCambio}). Gasto Real de ${mesTexto}: ${moneda2(json.gastoRealAnterior)} → ${moneda2(json.gastoRealNuevo)}.`,
      });
      setForm((prev) => ({ ...prev, montoSoles: "", numeroComprobante: "", comentario: "" }));
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
        <h2 className="text-lg font-semibold">Facturas OPEX</h2>
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Registra cada factura en Soles (sin IGV) — la app la convierte sola a USD con el tipo de cambio
          {" "}{TIPO_CAMBIO_POR_DEFECTO} y suma ese monto al Gasto Real del mes que elijas en Presupuesto 2026. La
          factura queda guardada en &quot;Facturas Opex - App&quot; (se crea sola la primera vez que la uses).
        </p>
      </div>

      <form onSubmit={registrar} className="card p-4 flex flex-col gap-3">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label className="etiqueta">Grupo de Gasto</label>
            <select className="campo" value={grupoSel} onChange={(e) => setGrupoSel(e.target.value)} required>
              <option value="">Selecciona…</option>
              {grupos.map((g) => (
                <option key={g} value={g}>
                  {g}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="etiqueta">Subgrupo de Gasto</label>
            <select
              className="campo"
              value={subgrupoSel}
              onChange={(e) => setSubgrupoSel(e.target.value)}
              disabled={!grupoSel}
              required
            >
              <option value="">{grupoSel ? "Selecciona…" : "Primero elige un grupo"}</option>
              {subgrupos.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="etiqueta">Línea de Gasto</label>
            <select
              className="campo"
              value={lineaSel}
              onChange={(e) => setLineaSel(e.target.value)}
              disabled={!subgrupoSel}
              required
            >
              <option value="">{subgrupoSel ? "Selecciona…" : "Primero elige un subgrupo"}</option>
              {lineasDisponibles.map((l) => (
                <option key={l.filaExcel} value={l.filaExcel}>
                  {l.lineaGasto}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="etiqueta">Mes (período real del gasto)</label>
            <select className="campo" value={form.mes} onChange={(e) => setForm((p) => ({ ...p, mes: e.target.value }))} required>
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  {nombre}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="etiqueta">Monto en Soles (sin IGV)</label>
            <input
              type="number"
              step="0.01"
              min="0"
              className="campo"
              value={form.montoSoles}
              onChange={(e) => setForm((p) => ({ ...p, montoSoles: e.target.value }))}
              required
            />
            {montoUsdPrevio != null && (
              <p className="text-xs mt-1" style={{ color: "var(--texto-suave)" }}>
                ≈ {moneda2(montoUsdPrevio)} al tipo de cambio {TIPO_CAMBIO_POR_DEFECTO}
              </p>
            )}
          </div>
          <div>
            <label className="etiqueta">Proveedor</label>
            <input
              type="text"
              className="campo"
              value={form.proveedor}
              onChange={(e) => setForm((p) => ({ ...p, proveedor: e.target.value }))}
              required
            />
          </div>

          <div>
            <label className="etiqueta">N° Comprobante</label>
            <input
              type="text"
              className="campo"
              value={form.numeroComprobante}
              onChange={(e) => setForm((p) => ({ ...p, numeroComprobante: e.target.value }))}
            />
          </div>
          <div className="sm:col-span-2">
            <label className="etiqueta">Comentario (opcional)</label>
            <input
              type="text"
              className="campo"
              value={form.comentario}
              onChange={(e) => setForm((p) => ({ ...p, comentario: e.target.value }))}
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

      <div className="card p-0 overflow-hidden">
        <div className="p-4 pb-0">
          <h3 className="font-semibold">Últimas facturas registradas</h3>
        </div>
        <div className="overflow-x-auto p-4">
          <table className="border-collapse" style={{ tableLayout: "fixed", width: "100%", minWidth: 1200 }}>
            <colgroup>
              <col style={{ width: 100 }} />
              <col style={{ width: 220 }} />
              <col style={{ width: 90 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 110 }} />
              <col style={{ width: 110 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 200 }} />
            </colgroup>
            <thead>
              <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                <th className="py-2 pr-3">Fecha</th>
                <th className="py-2 pr-3">Línea de Gasto</th>
                <th className="py-2 pr-3">Mes</th>
                <th className="py-2 pr-3">Proveedor</th>
                <th className="py-2 pr-3 text-right">Monto (USD)</th>
                <th className="py-2 pr-3 text-right">Soles (sin IGV)</th>
                <th className="py-2 pr-3">N° Comprobante</th>
                <th className="py-2 pr-3">Comentario</th>
              </tr>
            </thead>
            <tbody>
              {(datos?.facturas ?? []).slice(0, 25).map((f) => (
                <tr key={f.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td className="py-1.5 pr-3 text-xs" style={{ color: "var(--texto-suave)" }}>
                    {f.fecha}
                  </td>
                  <td className="py-1.5 pr-3 truncate" style={{ color: "var(--texto-suave)" }} title={f.lineaGasto}>
                    {f.lineaGasto}
                  </td>
                  <td className="py-1.5 pr-3 text-xs" style={{ color: "var(--texto-suave)" }}>
                    {f.mes ? NOMBRES_MES_CIERRE[f.mes - 1] : "—"}
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="proveedor"
                      tipo="texto"
                      valor={f.proveedor}
                      endpoint="/api/opex/facturas/editar-campo"
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { proveedor: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <MontoFactura factura={f} onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)} />
                  </td>
                  <td className="py-1.5 pr-3 text-right text-xs" style={{ color: "var(--texto-suave)" }}>
                    {f.montoSoles != null
                      ? `S/ ${f.montoSoles.toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}${f.tipoCambio ? ` (TC ${f.tipoCambio})` : ""}`
                      : "—"}
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="numeroComprobante"
                      tipo="texto"
                      valor={f.numeroComprobante}
                      endpoint="/api/opex/facturas/editar-campo"
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { numeroComprobante: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="comentario"
                      tipo="texto"
                      valor={f.comentario}
                      endpoint="/api/opex/facturas/editar-campo"
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { comentario: String(v) })}
                    />
                  </td>
                </tr>
              ))}
              {(datos?.facturas ?? []).length === 0 && (
                <tr>
                  <td colSpan={8} className="py-4 text-center text-sm" style={{ color: "var(--texto-suave)" }}>
                    Todavía no hay facturas registradas desde la app.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/** Edita el Monto (USD) — siempre resoluble aquí porque cada factura ya guarda su línea
 *  de gasto y mes de forma directa al registrarla (a diferencia de CAPEX, no hace falta
 *  adivinar por texto). Esta edición sigue siendo directo en USD — a propósito no se
 *  tocó, el cambio de Soles/IGV es solo para el registro inicial. */
function MontoFactura({ factura, onGuardado }: { factura: FacturaOpex; onGuardado: (cambios: Partial<FacturaOpex>) => void }) {
  const [valorLocal, setValorLocal] = useState(String(factura.monto));
  const [enfocado, setEnfocado] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!enfocado) setValorLocal(String(factura.monto));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [factura.monto]);

  if (!factura.filaPresupuesto || !factura.mes) {
    return (
      <span
        className="text-right block text-xs"
        style={{ color: "var(--texto-suave)" }}
        title="Esta factura no tiene guardada la línea de gasto o el mes, así que el monto no se puede corregir aquí."
      >
        {moneda2(factura.monto)} 🔒
      </span>
    );
  }

  async function guardar() {
    setEnfocado(false);
    const nuevo = Number(valorLocal);
    if (!Number.isFinite(nuevo) || nuevo <= 0 || nuevo === factura.monto) {
      setValorLocal(String(factura.monto));
      return;
    }
    setError(null);
    setGuardando(true);
    try {
      const res = await fetch("/api/opex/facturas/editar-monto", {
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
