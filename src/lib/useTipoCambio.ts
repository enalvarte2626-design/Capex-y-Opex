"use client";

import { useEffect, useState } from "react";
import { TIPO_CAMBIO_POR_DEFECTO } from "./opex-constantes";

// Re-exportado tal cual para no romper a nadie que ya importe esto desde acá — el valor
// real vive en opex-constantes.ts (sin "use client") para que las rutas de API del
// servidor también puedan importarlo de forma segura. Ver el comentario en ese archivo.
export { TIPO_CAMBIO_POR_DEFECTO };

const CLAVE_STORAGE = "capex-tipo-cambio";

/**
 * Tipo de cambio USD→PEN usado por la vista "en Soles" — editable desde la UI, se guarda
 * en localStorage para quedar igual entre el Dashboard y el Detalle BD_CAPEX. Nunca se
 * escribe en el Excel: es solo una conversión de referencia para mostrar en pantalla.
 */
export function useTipoCambio(): [number, (v: number) => void] {
  const [tipoCambio, setTipoCambio] = useState(TIPO_CAMBIO_POR_DEFECTO);

  useEffect(() => {
    const guardado = window.localStorage.getItem(CLAVE_STORAGE);
    if (guardado) {
      const n = parseFloat(guardado);
      if (n > 0) setTipoCambio(n);
    }
  }, []);

  function actualizar(v: number) {
    if (!(v > 0)) return;
    setTipoCambio(v);
    window.localStorage.setItem(CLAVE_STORAGE, String(v));
  }

  return [tipoCambio, actualizar];
}
