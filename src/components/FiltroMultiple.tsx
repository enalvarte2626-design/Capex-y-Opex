"use client";

import { useEffect, useRef, useState } from "react";

interface Opcion {
  valor: string;
  cantidad?: number;
}

interface Props {
  etiqueta: string;
  opciones: Opcion[];
  seleccion: string[];
  onCambiar: (v: string[]) => void;
}

/**
 * Filtro de selección múltiple en un menú desplegable (checklist), en vez de una fila
 * de chips — para listas largas (ej. Grupo de Gasto en OPEX) donde una fila de botones
 * se ve pesada. Muestra "Todos" cuando está todo seleccionado, o "N de M" si hay un
 * recorte.
 */
export default function FiltroMultiple({ etiqueta, opciones, seleccion, onCambiar }: Props) {
  const [abierto, setAbierto] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function alHacerClicFuera(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setAbierto(false);
    }
    document.addEventListener("mousedown", alHacerClicFuera);
    return () => document.removeEventListener("mousedown", alHacerClicFuera);
  }, []);

  function alternar(valor: string) {
    onCambiar(seleccion.includes(valor) ? seleccion.filter((v) => v !== valor) : [...seleccion, valor]);
  }

  const textoResumen =
    seleccion.length === 0
      ? "Ninguno"
      : seleccion.length === opciones.length
        ? "Todos"
        : `${seleccion.length} de ${opciones.length}`;

  return (
    <div className="relative" ref={ref}>
      <button
        type="button"
        className="campo flex items-center gap-2"
        style={{ width: "auto", cursor: "pointer" }}
        onClick={() => setAbierto((v) => !v)}
      >
        <span className="etiqueta mb-0" style={{ display: "inline" }}>
          {etiqueta}:
        </span>
        <span className="font-medium">{textoResumen}</span>
        <span style={{ color: "var(--texto-suave)" }}>{abierto ? "▲" : "▾"}</span>
      </button>

      {abierto && (
        <div
          className="card p-2"
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            zIndex: 20,
            minWidth: 260,
            maxHeight: 320,
            overflowY: "auto",
          }}
        >
          <div className="flex gap-2 px-1 pb-2 mb-1" style={{ borderBottom: "1px solid var(--borde)" }}>
            <button
              type="button"
              className="text-xs font-medium hover:underline"
              style={{ color: "var(--acento)" }}
              onClick={() => onCambiar(opciones.map((o) => o.valor))}
            >
              Seleccionar todos
            </button>
            <button
              type="button"
              className="text-xs font-medium hover:underline"
              style={{ color: "var(--texto-suave)" }}
              onClick={() => onCambiar([])}
            >
              Ninguno
            </button>
          </div>
          {opciones.map((o) => (
            <label
              key={o.valor}
              className="flex items-center gap-2 px-1 py-1 rounded text-sm cursor-pointer"
              style={{ userSelect: "none" }}
              onMouseDown={(e) => e.preventDefault()}
            >
              <input
                type="checkbox"
                checked={seleccion.includes(o.valor)}
                onChange={() => alternar(o.valor)}
              />
              <span>
                {o.valor}
                {o.cantidad != null && (
                  <span className="ml-1 text-xs" style={{ color: "var(--texto-suave)" }}>
                    ({o.cantidad})
                  </span>
                )}
              </span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}
