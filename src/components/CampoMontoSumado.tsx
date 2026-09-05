"use client";

import { useEffect, useState } from "react";
import { moneda2 } from "@/lib/format";

interface Props {
  fila: number;
  campo: string;
  valor: number;
  onGuardado: (nuevoValor: number) => void;
  className?: string;
  /** Endpoint a llamar (por defecto /api/capex/celda) — para reusar este mismo campo
   *  con otros módulos (ej. OPEX), manteniendo el cuerpo {fila, campo, valor}. */
  endpoint?: string;
  /** Si se pasan, muestra el equivalente en Soles debajo del campo (mismo criterio que
   *  las demás columnas de importe) — no afecta el guardado, es solo referencia visual. */
  mostrarSoles?: boolean;
  tipoCambio?: number;
}

/**
 * Campo de monto mensual (Real/Proyectado) con opción de sumar varios gastos: por
 * defecto edita el total directo, como cualquier campo numérico. El botón "Σ" abre una
 * lista donde se puede escribir cada gasto por separado (un proyecto puede tener más de
 * 3 en el mismo mes) — la app los suma y guarda solo el total en esa celda de Excel; los
 * montos individuales no quedan guardados en ningún lado, son solo una ayuda para sumar.
 */
export default function CampoMontoSumado({
  fila,
  campo,
  valor,
  onGuardado,
  className,
  endpoint = "/api/capex/celda",
  mostrarSoles = false,
  tipoCambio = 1,
}: Props) {
  const [valorLocal, setValorLocal] = useState(String(valor));
  const [enfocado, setEnfocado] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [expandido, setExpandido] = useState(false);
  const [lineas, setLineas] = useState<string[]>([]);

  useEffect(() => {
    if (!enfocado && !expandido) setValorLocal(String(valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [valor]);

  async function guardarValor(nuevoValor: number) {
    const valorOriginal = valor;
    if (nuevoValor === valorOriginal) return;
    setError(null);
    setGuardando(true);
    try {
      const res = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fila, campo, valor: nuevoValor }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
      onGuardado(nuevoValor);
    } catch (e) {
      setValorLocal(String(valorOriginal));
      setError((e as Error).message);
      throw e;
    } finally {
      setGuardando(false);
    }
  }

  async function guardarSimple() {
    setEnfocado(false);
    const texto = valorLocal.trim();
    const numero = Number(texto || 0);
    if (!Number.isFinite(numero)) {
      setValorLocal(String(valor));
      return;
    }
    await guardarValor(numero).catch(() => {});
  }

  function abrirSuma() {
    setLineas(valor ? [String(valor)] : [""]);
    setExpandido(true);
  }

  const totalLineas = lineas.reduce((a, l) => a + (Number(l) || 0), 0);

  async function guardarSuma() {
    await guardarValor(totalLineas).catch(() => {});
    setExpandido(false);
  }

  if (expandido) {
    return (
      <div className="rounded-md p-2" style={{ background: "var(--card)", border: "1px solid var(--acento)" }}>
        <div className="flex flex-col gap-1">
          {lineas.map((linea, i) => (
            <div key={i} className="flex items-center gap-1">
              <input
                type="number"
                autoFocus={i === lineas.length - 1}
                className="campo text-xs"
                style={{ padding: "0.2rem 0.4rem" }}
                value={linea}
                placeholder="0.00"
                onChange={(e) => setLineas((prev) => prev.map((l, j) => (j === i ? e.target.value : l)))}
              />
              <button
                type="button"
                className="text-xs px-1"
                style={{ color: "var(--texto-suave)" }}
                onClick={() => setLineas((prev) => prev.filter((_, j) => j !== i))}
                title="Quitar esta línea"
              >
                ×
              </button>
            </div>
          ))}
        </div>
        <button
          type="button"
          className="text-xs mt-1"
          style={{ color: "var(--acento)" }}
          onClick={() => setLineas((prev) => [...prev, ""])}
        >
          + Agregar gasto
        </button>
        <div className="text-xs font-semibold mt-1" style={{ borderTop: "1px solid var(--borde)", paddingTop: 4 }}>
          Total: {moneda2(totalLineas)}
        </div>
        <div className="flex gap-1 mt-1">
          <button type="button" className="boton-primario text-xs" style={{ padding: "0.25rem 0.6rem" }} onClick={guardarSuma} disabled={guardando}>
            Guardar
          </button>
          <button
            type="button"
            className="boton-secundario text-xs"
            style={{ padding: "0.25rem 0.6rem" }}
            onClick={() => setExpandido(false)}
            disabled={guardando}
          >
            Cancelar
          </button>
        </div>
        {error && (
          <p className="text-xs mt-0.5" style={{ color: "var(--peligro)" }}>
            {error}
          </p>
        )}
      </div>
    );
  }

  const valorMostrado = !enfocado ? moneda2(Number(valorLocal) || 0) : valorLocal;

  return (
    <div>
      <div className="flex items-center gap-0.5">
        <input
          type="text"
          inputMode="decimal"
          className={className}
          style={{
            background: "transparent",
            border: "1px solid transparent",
            width: "100%",
            opacity: guardando ? 0.6 : 1,
          }}
          value={valorMostrado}
          disabled={guardando}
          onChange={(e) => setValorLocal(e.target.value)}
          onBlur={guardarSimple}
          onKeyDown={(e) => {
            if (e.key === "Enter") (e.target as HTMLInputElement).blur();
          }}
          onFocus={(e) => {
            setEnfocado(true);
            e.target.style.border = "1px solid var(--acento)";
          }}
        />
        <button
          type="button"
          title="Sumar varios gastos de este mes"
          className="text-xs shrink-0"
          style={{ color: "var(--texto-suave)" }}
          onMouseDown={(e) => e.preventDefault()}
          onClick={abrirSuma}
        >
          Σ
        </button>
      </div>
      {mostrarSoles && !enfocado && (
        <div className="text-[10px] text-center" style={{ color: "var(--texto-suave)" }}>
          S/ {((Number(valorLocal) || 0) * tipoCambio).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </div>
      )}
      {error && (
        <p className="text-xs mt-0.5" style={{ color: "var(--peligro)" }}>
          {error}
        </p>
      )}
    </div>
  );
}
