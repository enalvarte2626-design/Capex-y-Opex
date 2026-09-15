"use client";

import { createContext, useContext } from "react";

interface Anios {
  capex: string;
  opex: string;
}

const AnioContext = createContext<Anios>({ capex: "2026", opex: "2026" });

interface Props extends Anios {
  children: React.ReactNode;
}

/** Reparte el año de presupuesto activo de CAPEX y de OPEX (calculados una sola vez en
 *  el servidor, en el layout, leyendo la hoja "Config App" de cada Excel) a toda la app
 *  de cliente — para que ningún componente tenga que volver a escribir "2026" a mano.
 *  Cada módulo tiene su propio año porque viven en archivos de Excel distintos. */
export function AnioProvider({ capex, opex, children }: Props) {
  return <AnioContext.Provider value={{ capex, opex }}>{children}</AnioContext.Provider>;
}

/** { capex, opex }: año activo de cada módulo (ej. "2026") para mostrar en títulos,
 *  menús y mensajes. Ver lib/anio.ts para cómo se calcula cada uno. */
export function useAnios(): Anios {
  return useContext(AnioContext);
}
