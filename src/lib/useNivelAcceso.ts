"use client";

import { useEffect, useState } from "react";

export type NivelAcceso = "completo" | "lectura" | "ninguno";

/**
 * Nivel de acceso de quien mira la página ahora mismo — para ocultar formularios y
 * campos editables a quien entró con la contraseña de solo lectura (o sin ninguna, en
 * páginas públicas como Facturas). Mientras no se sepa (recién montado), asume
 * "completo" para no hacer parpadear la pantalla — el servidor de todas formas rechaza
 * cualquier escritura que no corresponda, así que esto es solo para la interfaz, nunca
 * la única barrera real.
 */
export function useNivelAcceso(): NivelAcceso {
  const [nivel, setNivel] = useState<NivelAcceso>("completo");

  useEffect(() => {
    fetch("/api/auth/nivel", { cache: "no-store" })
      .then((r) => r.json())
      .then((j) => {
        if (j?.nivel === "completo" || j?.nivel === "lectura" || j?.nivel === "ninguno") {
          setNivel(j.nivel);
        }
      })
      .catch(() => {
        /* si falla, se queda en "completo" — el servidor igual valida cada escritura */
      });
  }, []);

  return nivel;
}
