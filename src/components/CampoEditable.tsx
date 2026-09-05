"use client";

import { useEffect, useState } from "react";
import { moneda2 } from "@/lib/format";

interface Props {
  fila: number;
  campo: string;
  valor: string | number;
  tipo: "texto" | "numero" | "moneda";
  /** Se llama tras guardar con éxito, para reflejar el nuevo valor en el estado local. */
  onGuardado: (nuevoValor: string | number) => void;
  className?: string;
  placeholder?: string;
  /** Endpoint a llamar (por defecto /api/capex/celda) — para reusar este mismo campo
   *  con otras hojas/rutas (ej. facturas), manteniendo el cuerpo {fila, campo, valor}. */
  endpoint?: string;
  /** Se llama antes de guardar (solo si el valor cambió) — si devuelve false, cancela el
   *  guardado y deja el valor como estaba. Para avisar de un efecto colateral (ej. romper
   *  el vínculo con facturas ya registradas) antes de escribir en el Excel. */
  confirmarAntes?: () => boolean;
}

/**
 * Input editable "in place": muestra el valor, guarda al perder el foco (o Enter) si
 * cambió, y escribe esa única celda en el Excel real vía /api/capex/celda. Optimista:
 * si falla, revierte al valor anterior y muestra el error debajo.
 *
 * tipo="moneda": muestra "$1,400.00" mientras no se edita, y el número plano (editable)
 * al enfocar — mismo patrón que un input de monto en cualquier formulario financiero.
 */
export default function CampoEditable({
  fila,
  campo,
  valor,
  tipo,
  onGuardado,
  className,
  placeholder,
  endpoint = "/api/capex/celda",
  confirmarAntes,
}: Props) {
  const [valorLocal, setValorLocal] = useState(String(valor));
  const [enfocado, setEnfocado] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Si el valor cambia desde afuera (recarga, cambio de mes de cierre) y no se está
  // editando ahora mismo, refleja el nuevo valor — evita mostrar algo desactualizado.
  useEffect(() => {
    if (!enfocado) setValorLocal(String(valor));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [valor]);

  async function guardar() {
    setEnfocado(false);
    const nuevoValorTexto = valorLocal.trim();
    const valorOriginalTexto = String(valor);
    if (nuevoValorTexto === valorOriginalTexto) return;

    if (confirmarAntes && !confirmarAntes()) {
      setValorLocal(valorOriginalTexto);
      return;
    }

    const esNumerico = tipo === "numero" || tipo === "moneda";
    const valorAEnviar = esNumerico ? Number(nuevoValorTexto || 0) : nuevoValorTexto;
    if (esNumerico && !Number.isFinite(valorAEnviar)) {
      setValorLocal(valorOriginalTexto);
      return;
    }

    setError(null);
    setGuardando(true);
    try {
      const res = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fila, campo, valor: valorAEnviar }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "No se pudo guardar.");
      onGuardado(valorAEnviar);
    } catch (e) {
      setValorLocal(valorOriginalTexto);
      setError((e as Error).message);
    } finally {
      setGuardando(false);
    }
  }

  const esMoneda = tipo === "moneda";
  const valorMostrado = esMoneda && !enfocado ? moneda2(Number(valorLocal) || 0) : valorLocal;

  return (
    <div>
      <input
        type={esMoneda ? "text" : tipo === "numero" ? "number" : "text"}
        inputMode={esMoneda ? "decimal" : undefined}
        className={className}
        style={{
          background: "transparent",
          border: "1px solid transparent",
          width: "100%",
          opacity: guardando ? 0.6 : 1,
        }}
        value={valorMostrado}
        placeholder={placeholder}
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
