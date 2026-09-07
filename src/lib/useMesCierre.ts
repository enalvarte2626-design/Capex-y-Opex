"use client";

import { useEffect, useState } from "react";
import { MES_CIERRE_POR_DEFECTO } from "./opex-constantes";

// Re-exportado tal cual para no romper a nadie que ya importe esto desde acá — el valor
// real vive en opex-constantes.ts (sin "use client") para que las rutas de API del
// servidor también puedan importarlo de forma segura. Ver el comentario en ese archivo.
export { MES_CIERRE_POR_DEFECTO };

const CLAVE_STORAGE = "capex-mes-cierre";

/**
 * Mes de cierre (1-12): el último mes con Gasto Real ya cerrado — de ahí en adelante se
 * cuenta como Forecast. Se guarda en localStorage para que quede igual al navegar entre
 * el Dashboard y el Detalle, pero nunca se escribe en el Excel — es solo una preferencia
 * local de cómo mirar los datos.
 */
export function useMesCierre(): [number, (v: number) => void] {
  const [mesCierre, setMesCierre] = useState(MES_CIERRE_POR_DEFECTO);

  useEffect(() => {
    const guardado = window.localStorage.getItem(CLAVE_STORAGE);
    if (guardado) {
      const n = parseInt(guardado, 10);
      if (n >= 1 && n <= 12) setMesCierre(n);
    }
  }, []);

  function actualizar(v: number) {
    setMesCierre(v);
    window.localStorage.setItem(CLAVE_STORAGE, String(v));
  }

  return [mesCierre, actualizar];
}
