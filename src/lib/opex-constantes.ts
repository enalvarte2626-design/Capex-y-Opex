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
