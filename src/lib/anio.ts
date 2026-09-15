/**
 * Año de presupuesto que la app está mostrando — un solo lugar para cambiarlo cada año
 * nuevo (2027, 2028…) SIN TOCAR CÓDIGO: basta con actualizar la variable de entorno
 * ANIO_PRESUPUESTO en Azure (App Settings del App Service) y reiniciar la app — Azure ya
 * reinicia solo al guardar un cambio de configuración, no hace falta un despliegue nuevo.
 * Sin esa variable, usa "2026" (el año con el que se armó la app) como valor por defecto.
 *
 * Sin "use client" a propósito: así tanto el layout (servidor, para el título de la
 * pestaña) como cualquier ruta de API pueden importarlo directo. Los componentes de
 * cliente (menús, títulos de pantalla) lo reciben en cambio vía `useAnio()` en
 * `AnioProvider.tsx` — nunca leen `process.env` ellos mismos, porque en un componente de
 * cliente esa lectura quedaría fija en lo que hubiera al compilar, no en lo que diga la
 * variable de entorno en producción.
 *
 * OJO: esto solo cambia el AÑO que se muestra en pantalla (títulos, menú). Si el Excel
 * real usa una hoja distinta cada año (ej. "Presupuesto 2026" → "Presupuesto 2027"), hay
 * que actualizar también las variables SP_OPEX_HOJA_PRESUPUESTO / SP_CAPEX_HOJA_PROYECCION
 * — ambas ya son variables de entorno, así que tampoco necesitan tocar código.
 */
export function anioPresupuestoActual(): string {
  return process.env.ANIO_PRESUPUESTO?.trim() || "2026";
}
