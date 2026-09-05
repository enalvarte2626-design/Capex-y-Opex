"use client";

import { useEffect, useMemo, useState } from "react";
import { NOMBRES_MES_CIERRE, type FacturaCapex } from "@/lib/capex";
import { moneda2 } from "@/lib/format";
import CampoEditable from "@/components/CampoEditable";

interface ProyectoOpcion {
  filaExcel: number;
  proyecto: string;
  detalle: string;
  grupoNegocio: string;
  responsable: string;
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
  const [datos, setDatos] = useState<Respuesta | null>(null);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [proyectoNombreSel, setProyectoNombreSel] = useState("");
  const [detalleSel, setDetalleSel] = useState("");

  const [form, setForm] = useState({
    filaProyecto: "",
    mes: String(new Date().getMonth() + 1),
    monto: "",
    recurso: "",
    proveedor: "",
    responsable: "",
    numeroFactura: "",
    periodoFacturado: HOY(),
    comentarioExtra: "",
  });
  const [guardando, setGuardando] = useState(false);
  const [mensaje, setMensaje] = useState<{ tipo: "ok" | "error"; texto: string } | null>(null);

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

  const nombresProyecto = useMemo(() => {
    const conteo = new Map<string, number>();
    for (const p of datos?.proyectos ?? []) conteo.set(p.proyecto, (conteo.get(p.proyecto) ?? 0) + 1);
    return Array.from(conteo.entries())
      .sort((a, b) => a[0].localeCompare(b[0], "es"))
      .map(([nombre, cantidad]) => ({ nombre, cantidad }));
  }, [datos]);

  // El desplegable de Detalle se acota solo al Proyecto ya elegido.
  const detallesDisponibles = useMemo(() => {
    const lista = (datos?.proyectos ?? []).filter((p) => !proyectoNombreSel || p.proyecto === proyectoNombreSel);
    const vistos = new Set<string>();
    const resultado: { nombre: string; filaExcel: number }[] = [];
    for (const p of lista) {
      const d = p.detalle.trim() || "(sin detalle)";
      if (vistos.has(d)) continue;
      vistos.add(d);
      resultado.push({ nombre: d, filaExcel: p.filaExcel });
    }
    return resultado.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  }, [datos, proyectoNombreSel]);

  // Al cambiar el Proyecto, el Detalle elegido antes puede ya no aplicar.
  useEffect(() => {
    setDetalleSel("");
  }, [proyectoNombreSel]);

  // En cuanto Proyecto + Detalle identifican una sola fila de BD_CAPEX, la deja lista
  // (y sugiere su Responsable, si el campo aún está vacío).
  useEffect(() => {
    if (!detalleSel) return;
    const opcion = detallesDisponibles.find((d) => d.nombre === detalleSel);
    if (!opcion) return;
    const p = datos?.proyectos.find((x) => x.filaExcel === opcion.filaExcel);
    setForm((prev) => ({
      ...prev,
      filaProyecto: String(opcion.filaExcel),
      responsable: prev.responsable || p?.responsable || "",
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [detalleSel]);

  const proyectoElegido = datos?.proyectos.find((p) => String(p.filaExcel) === form.filaProyecto);

  function actualizarCampo(campo: keyof typeof form, valor: string) {
    setForm((prev) => {
      const siguiente = { ...prev, [campo]: valor };
      // Al elegir proyecto, sugiere su Responsable si el campo aún está vacío.
      if (campo === "filaProyecto") {
        const p = datos?.proyectos.find((x) => String(x.filaExcel) === valor);
        if (p && !prev.responsable) siguiente.responsable = p.responsable;
      }
      return siguiente;
    });
  }

  async function registrar(e: React.FormEvent) {
    e.preventDefault();
    if (!proyectoElegido) {
      setMensaje({ tipo: "error", texto: "Elige un proyecto." });
      return;
    }
    const monto = Number(form.monto);
    if (!Number.isFinite(monto) || monto <= 0) {
      setMensaje({ tipo: "error", texto: "El monto debe ser mayor a 0." });
      return;
    }

    const mesTexto = NOMBRES_MES_CIERRE[Number(form.mes) - 1];
    const confirmado = window.confirm(
      `¿Registrar factura de ${moneda2(monto)} para "${proyectoElegido.proyecto} — ${proyectoElegido.detalle || "(sin detalle)"}", período ${mesTexto}? Esto suma ${moneda2(monto)} al Gasto Real de ${mesTexto} en BD_CAPEX y agrega una fila en la hoja de facturas.`
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
          monto,
          recurso: form.recurso,
          proveedor: form.proveedor,
          responsable: form.responsable,
          numeroFactura: form.numeroFactura,
          periodoFacturado: form.periodoFacturado,
          comentarioExtra: form.comentarioExtra || undefined,
        }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo registrar la factura.");
      setMensaje({ tipo: "ok", texto: `Factura registrada. Gasto Real de ${mesTexto}: ${moneda2(json.gastoRealAnterior)} → ${moneda2(json.gastoRealNuevo)}.` });
      setForm((prev) => ({
        ...prev,
        monto: "",
        numeroFactura: "",
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
        <p className="text-sm" style={{ color: "var(--texto-suave)" }}>
          Registra cada factura de un proyecto — el monto se suma solo al Gasto Real del mes que elijas en
          BD_CAPEX, y la factura queda guardada en &quot;Control de Facturas-Capex 25fEB&quot;.
        </p>
      </div>

      <form onSubmit={registrar} className="card p-4 flex flex-col gap-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="etiqueta">Proyecto</label>
            <select
              className="campo"
              value={proyectoNombreSel}
              onChange={(e) => setProyectoNombreSel(e.target.value)}
              required
            >
              <option value="">Selecciona un proyecto…</option>
              {nombresProyecto.map((p) => (
                <option key={p.nombre} value={p.nombre}>
                  {p.nombre} ({p.cantidad})
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="etiqueta">Detalle</label>
            <select
              className="campo"
              value={detalleSel}
              onChange={(e) => setDetalleSel(e.target.value)}
              disabled={!proyectoNombreSel}
              required
            >
              <option value="">{proyectoNombreSel ? "Selecciona un detalle…" : "Primero elige un proyecto"}</option>
              {detallesDisponibles.map((d) => (
                <option key={d.filaExcel} value={d.nombre}>
                  {d.nombre}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="etiqueta">Mes (período real del gasto)</label>
            <select className="campo" value={form.mes} onChange={(e) => actualizarCampo("mes", e.target.value)} required>
              {NOMBRES_MES_CIERRE.map((nombre, i) => (
                <option key={nombre} value={i + 1}>
                  {nombre}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="etiqueta">Monto Final (sin IGV)</label>
            <input
              type="number"
              step="0.01"
              min="0"
              className="campo"
              value={form.monto}
              onChange={(e) => actualizarCampo("monto", e.target.value)}
              required
            />
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
            <label className="etiqueta">Proveedor</label>
            <input type="text" className="campo" value={form.recurso} onChange={(e) => actualizarCampo("recurso", e.target.value)} />
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
            <label className="etiqueta">Periodo facturado (fecha de la factura)</label>
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

      <div className="card p-0 overflow-hidden">
        <div className="p-4 pb-0">
          <h3 className="font-semibold">Últimas facturas registradas</h3>
        </div>
        <p className="px-4 text-xs" style={{ color: "var(--texto-suave)" }}>
          Todos los campos son editables directo en la tabla — se guardan al salir del campo. El Monto solo se
          puede corregir cuando la app identifica con certeza a qué fila/mes de BD_CAPEX corresponde (si no, sale
          de solo lectura, con una nota).
        </p>
        <div className="overflow-x-auto p-4">
          <table className="border-collapse" style={{ tableLayout: "fixed", width: "100%", minWidth: 1200 }}>
            <colgroup>
              <col style={{ width: 130 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 100 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 240 }} />
              <col style={{ width: 120 }} />
              <col style={{ width: 130 }} />
              <col style={{ width: 220 }} />
            </colgroup>
            <thead>
              <tr className="text-left" style={{ color: "var(--texto-suave)" }}>
                <th className="py-2 pr-3">Periodo facturado</th>
                <th className="py-2 pr-3">Proveedor</th>
                <th className="py-2 pr-3">Empresa</th>
                <th className="py-2 pr-3">Responsable</th>
                <th className="py-2 pr-3">Proyecto</th>
                <th className="py-2 pr-3 text-right">Monto</th>
                <th className="py-2 pr-3">N° Factura</th>
                <th className="py-2 pr-3">Comentarios</th>
              </tr>
            </thead>
            <tbody>
              {(datos?.facturas ?? []).slice(0, 25).map((f) => (
                <tr key={f.filaExcel} style={{ borderTop: "1px solid var(--borde)" }}>
                  <td className="py-1.5 pr-3">
                    <CampoFecha
                      factura={f}
                      onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="recurso"
                      tipo="texto"
                      valor={f.recurso}
                      endpoint="/api/facturas/editar-campo"
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
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { responsable: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3 truncate" style={{ color: "var(--texto-suave)" }} title={f.proyecto}>
                    {f.proyecto}
                  </td>
                  <td className="py-1.5 pr-3">
                    <MontoFactura factura={f} onGuardado={(cambios) => actualizarFacturaLocal(f.filaExcel, cambios)} />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="numeroFactura"
                      tipo="texto"
                      valor={f.numeroFactura}
                      endpoint="/api/facturas/editar-campo"
                      onGuardado={(v) => actualizarFacturaLocal(f.filaExcel, { numeroFactura: String(v) })}
                    />
                  </td>
                  <td className="py-1.5 pr-3">
                    <CampoEditable
                      fila={f.filaExcel}
                      campo="comentarios"
                      tipo="texto"
                      valor={f.comentarios}
                      endpoint="/api/facturas/editar-campo"
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
}: {
  factura: FacturaConResolucion;
  onGuardado: (cambios: Partial<FacturaConResolucion>) => void;
}) {
  const [valor, setValor] = useState(factura.periodoFacturadoISO);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => setValor(factura.periodoFacturadoISO), [factura.periodoFacturadoISO]);

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
 * Edita el Monto de una factura ya registrada. Solo permite guardar cuando la app puede
 * identificar con certeza a qué fila/mes de BD_CAPEX corresponde (mismo texto que se
 * escribió al registrarla) — si no, queda de solo lectura, porque corregir el monto sin
 * eso significaría no poder ajustar el Gasto Real, o arriesgar a tocar la celda
 * equivocada.
 */
function MontoFactura({
  factura,
  onGuardado,
}: {
  factura: FacturaConResolucion;
  onGuardado: (cambios: Partial<FacturaConResolucion>) => void;
}) {
  const [valorLocal, setValorLocal] = useState(String(factura.monto));
  const [enfocado, setEnfocado] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!enfocado) setValorLocal(String(factura.monto));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [factura.monto]);

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
    if (!Number.isFinite(nuevo) || nuevo <= 0 || nuevo === factura.monto) {
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