"use client";

import { createContext, useContext } from "react";

const AnioContext = createContext("2026");

interface Props {
  anio: string;
  children: React.ReactNode;
}

/** Reparte el año de presupuesto (calculado una sola vez en el servidor, en el layout,
 *  a partir de la variable de entorno ANIO_PRESUPUESTO) a toda la app de cliente — para
 *  que ningún componente tenga que volver a escribir "2026" a mano. */
export function AnioProvider({ anio, children }: Props) {
  return <AnioContext.Provider value={anio}>{children}</AnioContext.Provider>;
}

/** Año de presupuesto actual (ej. "2026") para mostrar en títulos, menús y mensajes.
 *  Ver el comentario de `anioPresupuestoActual` en lib/anio.ts para cómo cambiarlo. */
export function useAnio(): string {
  return useContext(AnioContext);
}
