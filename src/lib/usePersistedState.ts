"use client";

import { useEffect, useRef, useState } from "react";

/**
 * Como useState, pero recuerda el valor en localStorage bajo `clave` — para que un filtro
 * (Grupo, Prioridad, Avance, Proyecto, Habilitar en Soles, etc.) quede exactamente como lo
 * dejaste al navegar a otro módulo y volver, en vez de resetearse al valor por defecto.
 * Mismo criterio que ya usan useMesCierre/useTipoCambio, generalizado para cualquier valor
 * serializable en JSON (string, number, boolean, array, objeto).
 *
 * El valor por defecto se usa mientras se lee localStorage en el primer render (evita el
 * "flash" de un valor y luego otro) y también si no hay nada guardado todavía o el valor
 * guardado no es JSON válido.
 */
type ActualizarValor<T> = T | ((prev: T) => T);

/** Mismo perfil que el setter de useState — acepta tanto un valor directo (`set(x)`) como
 *  una función de actualización (`set((prev) => ...)`), para poder ser un reemplazo 1:1
 *  de useState en cualquier filtro existente sin tocar el resto del código que lo usa. */
export function usePersistedState<T>(clave: string, valorPorDefecto: T): [T, (v: ActualizarValor<T>) => void] {
  const [valor, setValor] = useState<T>(valorPorDefecto);
  const cargado = useRef(false);

  useEffect(() => {
    try {
      const guardado = window.localStorage.getItem(clave);
      if (guardado != null) setValor(JSON.parse(guardado) as T);
    } catch {
      // Valor corrupto o localStorage no disponible — se queda con el valor por defecto.
    } finally {
      cargado.current = true;
    }
    // Solo al montar: `clave` no cambia entre renders para un mismo filtro.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function actualizar(v: ActualizarValor<T>) {
    setValor((prev) => {
      const siguiente = typeof v === "function" ? (v as (prev: T) => T)(prev) : v;
      try {
        window.localStorage.setItem(clave, JSON.stringify(siguiente));
      } catch {
        // Si falla (localStorage lleno/deshabilitado), el filtro sigue funcionando en memoria.
      }
      return siguiente;
    });
  }

  return [valor, actualizar];
}
