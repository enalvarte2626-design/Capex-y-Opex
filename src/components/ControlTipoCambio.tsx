"use client";

import { useEffect, useState } from "react";

interface Props {
  tipoCambio: number;
  onCambiar: (v: number) => void;
}

/**
 * Control junto al chip "Habilitar en Soles": muestra el tipo de cambio actual en un
 * campo editable y solo lo aplica (y recalcula todos los montos en S/ que dependen de
 * él) al presionar "Recalcular" — no en cada tecla, para no recalcular de más mientras
 * se está escribiendo el número nuevo.
 */
export default function ControlTipoCambio({ tipoCambio, onCambiar }: Props) {
  const [valor, setValor] = useState(String(tipoCambio));

  // Si el tipo de cambio cambia desde afuera (ej. se cargó desde localStorage al entrar
  // a esta página), refleja ese valor en el campo.
  useEffect(() => {
    setValor(String(tipoCambio));
  }, [tipoCambio]);

  function recalcular() {
    const n = parseFloat(valor.replace(",", "."));
    if (n > 0) onCambiar(n);
    else setValor(String(tipoCambio));
  }

  return (
    <div className="flex items-center gap-1">
      <span className="etiqueta mb-0">T.C. S/:</span>
      <input
        type="text"
        inputMode="decimal"
        className="campo"
        style={{ width: 70 }}
        value={valor}
        onChange={(e) => setValor(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") recalcular();
        }}
        title="Tipo de cambio USD→PEN para la vista en Soles (solo referencia)"
      />
      <button className="boton-secundario" onClick={recalcular} title="Vuelve a calcular todos los montos en Soles con este tipo de cambio">
        Recalcular
      </button>
    </div>
  );
}
