/**
 * Tipo de cambio por defecto USD↔PEN — módulo compartido SIN "use client".
 *
 * Antes esta constante vivía en `useTipoCambio.ts` (marcado "use client" porque ese
 * archivo también exporta el hook `useTipoCambio`). Al importarla desde ahí en rutas de
 * API (server-side), Next.js la reemplazaba por una "client reference" — un objeto
 * proxy que revienta si se llama como función, pero que como valor simplemente no es el
 * número 3.4: es un objeto. Eso hacía que `montoSoles / tipoCambio` diera `NaN`, y al
 * guardarlo en el Excel (`JSON.stringify(NaN) === "null"`) la celda de Monto (USD)
 * quedaba en blanco/0 — el bug de las facturas OPEX registradas con $0.00 pese a tener
 * el monto en Soles correcto.
 *
 * Por eso este valor vive en su propio archivo, sin "use client", para que tanto el
 * hook (cliente) como las rutas de API (servidor) lo importen del mismo lugar sin
 * cruzar el límite de compilación de Next.
 */
export const TIPO_CAMBIO_POR_DEFECTO = 3.4;

/** Tipo de cambio por defecto EUR→USD — mismo criterio y mismo motivo que
 *  `TIPO_CAMBIO_POR_DEFECTO` (sin "use client", para que server y cliente lo importen
 *  del mismo lugar sin cruzar el límite de compilación de Next). 1 EUR ≈ 1.10 USD. */
export const TIPO_CAMBIO_EUR_POR_DEFECTO = 1.1;

/**
 * Mes de cierre (1-12): el último mes con Gasto Real ya cerrado de verdad — de ahí en
 * adelante (mes de cierre + 1 en adelante) todavía está abierto y una factura nueva SÍ
 * debe sumar al Gasto Real de Presupuesto 2026 automáticamente al registrarla.
 *
 * Hoy en 7 (Julio) → Agosto (8) en adelante se considera abierto. Cuando se cierre
 * formalmente un mes más (ej. se cierre Agosto), subir este número a 8, y así sucesivamente.
 *
 * Vive acá (sin "use client") por el mismo motivo que TIPO_CAMBIO_POR_DEFECTO: se
 * necesita tanto en `useMesCierre.ts` (cliente, para el Dashboard) como en las rutas de
 * API de registro de facturas (servidor) — importarlo desde un archivo "use client" en
 * el servidor rompe silenciosamente.
 */
export const MES_CIERRE_POR_DEFECTO = 7;
