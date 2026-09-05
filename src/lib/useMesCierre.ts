"use client";

import { useEffect, useState } from "react";

/** Julio (índice 7 = mes 8 en 1-based cuenta desde 1)… en realidad: 1-12, "7" = Julio. */
export const MES_CIERRE_POR_DEFECTO = 7;

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
