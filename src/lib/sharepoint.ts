export interface ConfiguracionSharePoint {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  /** Enlace "compartir" del archivo, tal cual se copia desde SharePoint/Teams/OneDrive. */
  shareUrl: string;
}

const CAMPOS_REQUERIDOS: Array<keyof ConfiguracionSharePoint> = [
  "tenantId",
  "clientId",
  "clientSecret",
  "shareUrl",
];

const ETIQUETAS: Record<keyof ConfiguracionSharePoint, string> = {
  tenantId: "SP_TENANT_ID",
  clientId: "SP_CLIENT_ID",
  clientSecret: "SP_CLIENT_SECRET",
  shareUrl: "SP_CAPEX_SHARE_URL",
};

export function obtenerConfiguracionSharePoint(): ConfiguracionSharePoint {
  return {
    tenantId: process.env.SP_TENANT_ID ?? "",
    clientId: process.env.SP_CLIENT_ID ?? "",
    clientSecret: process.env.SP_CLIENT_SECRET ?? "",
    shareUrl: process.env.SP_CAPEX_SHARE_URL ?? "",
  };
}

export function camposFaltantes(config: ConfiguracionSharePoint): string[] {
  return CAMPOS_REQUERIDOS.filter((clave) => !config[clave]?.trim()).map((clave) => ETIQUETAS[clave]);
}

/** Misma cuenta de Entra ID (tenant/client) que CAPEX, apuntando al archivo de OPEX. */
export function obtenerConfiguracionOpex(): ConfiguracionSharePoint {
  return {
    tenantId: process.env.SP_TENANT_ID ?? "",
    clientId: process.env.SP_CLIENT_ID ?? "",
    clientSecret: process.env.SP_CLIENT_SECRET ?? "",
    shareUrl: process.env.SP_OPEX_SHARE_URL ?? "",
  };
}

export function camposFaltantesOpex(config: ConfiguracionSharePoint): string[] {
  const base = CAMPOS_REQUERIDOS.filter((clave) => !config[clave]?.trim()).map((clave) =>
    clave === "shareUrl" ? "SP_OPEX_SHARE_URL" : ETIQUETAS[clave]
  );
  return base;
}

/** Error con mensaje pensado para mostrarse directamente al usuario. */
export class ErrorSharePoint extends Error {}

interface TokenCache {
  token: string;
  expiraEn: number;
}
let cacheToken: TokenCache | null = null;

async function obtenerToken(config: ConfiguracionSharePoint): Promise<string> {
  if (cacheToken && cacheToken.expiraEn > Date.now() + 30_000) {
    return cacheToken.token;
  }

  const url = `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  let res: Response;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });
  } catch (e) {
    throw new ErrorSharePoint(
      `No se pudo contactar a Microsoft Entra ID (revisa tu conexión a internet y el Tenant ID). Detalle: ${(e as Error).message}`
    );
  }

  const datos = await res.json().catch(() => ({}));
  if (!res.ok) {
    const detalle = datos.error_description || datos.error || res.statusText;
    throw new ErrorSharePoint(
      `No se pudo autenticar en Microsoft Entra ID. Revisa Tenant ID, Client ID y Client Secret. Detalle: ${detalle}`
    );
  }

  cacheToken = { token: datos.access_token, expiraEn: Date.now() + datos.expires_in * 1000 };
  return cacheToken.token;
}

export async function graphFetch(
  config: ConfiguracionSharePoint,
  rutaOUrl: string,
  init: RequestInit = {}
): Promise<Response> {
  const token = await obtenerToken(config);
  const url = rutaOUrl.startsWith("https://") ? rutaOUrl : `https://graph.microsoft.com/v1.0${rutaOUrl}`;
  const res = await fetch(url, {
    // "no-store" a propósito: nunca se debe servir una respuesta de Graph desde el Data
    // Cache de Next.js — esto siempre tiene que ser el estado más reciente del Excel
    // (Presupuesto/Gasto real cambian a cada rato), nunca una copia vieja.
    cache: "no-store",
    ...init,
    headers: { ...init.headers, Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const cuerpo = await res.text().catch(() => "");
    let detalle = cuerpo;
    try {
      detalle = JSON.parse(cuerpo)?.error?.message || cuerpo;
    } catch {
      /* deja el texto crudo */
    }
    throw new ErrorSharePoint(`Microsoft Graph respondió ${res.status}: ${detalle}`);
  }
  return res;
}

/**
 * Convierte un enlace para compartir de SharePoint/OneDrive al identificador que espera
 * el endpoint /shares de Graph. Ver:
 * https://learn.microsoft.com/graph/api/shares-get
 */
function codificarShareUrl(url: string): string {
  const base64 = Buffer.from(url, "utf-8").toString("base64");
  const base64Url = base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return `u!${base64Url}`;
}

export interface ArchivoResuelto {
  driveId: string;
  itemId: string;
  nombre: string;
  /** Id de la carpeta que contiene el archivo — para crear archivos nuevos al lado. */
  carpetaId: string;
}

/**
 * Resuelve el enlace "compartir" directamente al archivo (drive + item id), sin
 * necesidad de conocer en qué carpeta vive dentro de SharePoint.
 */
export async function resolverArchivoPorShareUrl(
  config: ConfiguracionSharePoint
): Promise<ArchivoResuelto> {
  const shareId = codificarShareUrl(config.shareUrl.trim());
  const res = await graphFetch(
    config,
    `/shares/${shareId}/driveItem?$select=id,name,parentReference`
  );
  const datos = await res.json();
  return {
    driveId: datos.parentReference?.driveId as string,
    itemId: datos.id as string,
    nombre: datos.name as string,
    carpetaId: datos.parentReference?.id as string,
  };
}

/** Nombres de los archivos que ya existen en la misma carpeta (para evitar duplicados). */
export async function listarNombresCarpeta(
  config: ConfiguracionSharePoint,
  driveId: string,
  carpetaId: string
): Promise<string[]> {
  const res = await graphFetch(config, `/drives/${driveId}/items/${carpetaId}/children?$select=name`);
  const datos = await res.json();
  const items: Array<{ name: string }> = datos.value ?? [];
  return items.map((i) => i.name);
}

/** Sube un archivo nuevo (≤4 MB) a una carpeta puntual. Falla si ya existe uno con ese nombre. */
export async function crearArchivo(
  config: ConfiguracionSharePoint,
  driveId: string,
  carpetaId: string,
  nombre: string,
  contenido: Buffer
): Promise<ArchivoResuelto> {
  const ruta = `/drives/${driveId}/items/${carpetaId}:/${encodeURIComponent(nombre)}:/content`;
  const res = await graphFetch(config, `${ruta}?@microsoft.graph.conflictBehavior=fail`, {
    method: "PUT",
    headers: { "Content-Type": "application/octet-stream" },
    body: new Uint8Array(contenido),
  });
  const datos = await res.json();
  return {
    driveId,
    itemId: datos.id as string,
    nombre: datos.name as string,
    carpetaId,
  };
}

/** Descarga el contenido binario del archivo ya resuelto. */
export async function descargarContenido(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto
): Promise<Buffer> {
  const res = await graphFetch(config, `/drives/${archivo.driveId}/items/${archivo.itemId}/content`);
  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

/**
 * Escribe un único valor en una celda puntual del Excel en vivo (vía la API de
 * Workbook de Graph — el archivo no se descarga ni se vuelve a subir entero, solo se
 * actualiza esa celda). Pensado para ediciones puntuales desde la app, nunca en bloque.
 */
export async function escribirCelda(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  direccion: string,
  valor: number | string
): Promise<void> {
  const ruta = `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
    hoja
  )}')/range(address='${direccion}')`;
  await graphFetch(config, ruta, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ values: [[valor]] }),
  });
}

/** Lee el valor actual de una celda puntual (para sumarle algo antes de volver a escribir). */
export async function leerCelda(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  direccion: string
): Promise<number> {
  const ruta = `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
    hoja
  )}')/range(address='${direccion}')?$select=values`;
  const res = await graphFetch(config, ruta);
  const datos = await res.json();
  const valor = datos.values?.[0]?.[0];
  return typeof valor === "number" ? valor : Number(valor) || 0;
}

/** Escribe una fila completa (varias columnas seguidas) de una sola vez — para agregar una factura nueva. */
export async function escribirFila(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  fila: number,
  columnaInicio: string,
  columnaFin: string,
  valores: Array<string | number>
): Promise<void> {
  const direccion = `${columnaInicio}${fila}:${columnaFin}${fila}`;
  const ruta = `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
    hoja
  )}')/range(address='${direccion}')`;
  await graphFetch(config, ruta, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ values: [valores] }),
  });
}

/**
 * Inserta una fila en blanco en la posición indicada — igual que "Insertar fila" en
 * Excel: esa fila y todo lo que está debajo se corre un lugar hacia abajo, y Excel ajusta
 * solo las fórmulas que apuntaban a las filas desplazadas. Se usa cuando ya no queda
 * ningún hueco vacío para un registro nuevo (ej. un proyecto nuevo en BD_CAPEX), en vez
 * de fallar por falta de espacio. La dirección es la fila completa ("84:84"), no solo una
 * columna — así se corren todas las columnas juntas, no solo una.
 */
export async function insertarFila(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  fila: number
): Promise<void> {
  const ruta = `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
    hoja
  )}')/range(address='${fila}:${fila}')/insert`;
  await graphFetch(config, ruta, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ shift: "Down" }),
  });
}

/** Nombres de todas las hojas que ya tiene el archivo. */
export async function listarHojas(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto
): Promise<string[]> {
  const res = await graphFetch(
    config,
    `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets?$select=name`
  );
  const datos = await res.json();
  const hojas: Array<{ name: string }> = datos.value ?? [];
  return hojas.map((h) => h.name);
}

/** Crea una hoja nueva (al final del libro) si todavía no existe una con ese nombre. */
export async function crearHojaSiNoExiste(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  nombreHoja: string
): Promise<void> {
  const existentes = await listarHojas(config, archivo);
  if (existentes.some((n) => n.toLowerCase() === nombreHoja.toLowerCase())) return;
  await graphFetch(config, `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets/add`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: nombreHoja }),
  });
}

/**
 * Id + dirección (ej. "Hoja1!A1:O5") de la primera tabla de Excel de una hoja, o null.
 * `address` no es una propiedad de WorkbookTable ni se puede traer con `$expand=range`
 * (esa ruta tampoco es expandible en esta versión de Graph — ambas dan 400). Hay que
 * pedirla aparte, con una llamada directa a `tables('{id}')/range`.
 */
async function primeraTablaDeHoja(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string
): Promise<{ id: string; address: string } | null> {
  const res = await graphFetch(
    config,
    `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(hoja)}')/tables?$select=id`
  );
  const datos = await res.json();
  const tablas: Array<{ id: string }> = datos.value ?? [];
  const primera = tablas[0];
  if (!primera) return null;

  const resRango = await graphFetch(
    config,
    `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/tables('${primera.id}')/range?$select=address`
  );
  const rango = await resRango.json();
  return { id: primera.id, address: rango.address ?? "" };
}

/** "A" → 0, "B" → 1, ... "AA" → 26, igual que columnaALetra pero al revés. */
function letraAColumna(letra: string): number {
  let n = 0;
  for (const c of letra) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

/** Parsea una dirección tipo "'Hoja 1'!A1:O5" o "Hoja1!A1:O5" → sus 4 partes. */
function parsearRango(direccion: string): { colInicio: string; filaInicio: number; colFin: string; filaFin: number } | null {
  const sinHoja = direccion.includes("!") ? direccion.slice(direccion.lastIndexOf("!") + 1) : direccion;
  const m = sinHoja.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  return { colInicio: m[1], filaInicio: Number(m[2]), colFin: m[3], filaFin: Number(m[4]) };
}

/**
 * Asegura que la hoja tenga una Tabla de Excel real cubriendo AL MENOS el encabezado
 * indicado (ej. "A1:P1") — la crea si todavía no existe ninguna, y si ya existe pero es
 * más angosta (porque se agregó una columna nueva al esquema después de crearla, como
 * pasó con "Empresa" y luego "Moneda ingresada"), la ENSANCHA con `tables/{id}/resize`
 * para que incluya las columnas nuevas. Sin este ensanche, `agregarFilaTabla` falla con
 * "Microsoft Graph respondió 400: El número de filas o columnas... no coincide" en
 * cuanto se agrega una columna al esquema sin que la tabla ya creada en el Excel se entere.
 */
export async function asegurarTablaEnHoja(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  rangoEncabezado: string
): Promise<void> {
  const existente = await primeraTablaDeHoja(config, archivo, hoja);
  if (!existente) {
    await graphFetch(
      config,
      `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
        hoja
      )}')/tables/add`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ address: rangoEncabezado, hasHeaders: true }),
      }
    );
    return;
  }

  const rangoDeseado = parsearRango(rangoEncabezado);
  const rangoActual = parsearRango(existente.address);
  if (!rangoDeseado || !rangoActual) return; // no debería pasar; si pasa, sigue con la tabla tal cual está

  if (letraAColumna(rangoActual.colFin) >= letraAColumna(rangoDeseado.colFin)) return; // ya es suficientemente ancha

  // Ensancha manteniendo el mismo inicio y todas las filas que ya tenga la tabla —
  // resize() exige que el nuevo rango se solape con el actual y conserve el encabezado.
  const nuevaDireccion = `${rangoActual.colInicio}${rangoActual.filaInicio}:${rangoDeseado.colFin}${rangoActual.filaFin}`;
  await graphFetch(
    config,
    `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/tables('${existente.id}')/resize`,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ address: nuevaDireccion }),
    }
  );
}

/**
 * Agrega una fila al FINAL de la tabla de Excel de una hoja — usando el endpoint nativo
 * `tables/{id}/rows/add` de Graph en vez de calcular nosotros mismos "cuál es la última
 * fila con datos" y escribir ahí con `escribirFila`.
 *
 * Por qué: calcular la última fila nosotros (descargar el archivo entero, escanear,
 * sumar 1) tiene una condición de carrera real — si dos facturas se registran seguidas
 * (aunque sea con pocos segundos de diferencia, por la latencia de guardado/replicación
 * de SharePoint), ambas pueden leer el mismo "última fila" y la segunda escritura
 * termina PISANDO a la primera en la misma celda, sin ningún error visible: la app
 * responde "factura registrada" para ambas, pero solo una queda en el Excel. Esto es
 * justo lo que le pasó al usuario con una factura que el mensaje de éxito confirmó pero
 * que luego no aparecía en la lista. El endpoint de Tablas de Excel resuelve "agregar al
 * final" del lado del servidor de Graph, sin que nosotros tengamos que adivinar el
 * número de fila — así se elimina esa ventana de carrera.
 */
export async function agregarFilaTabla(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  valores: Array<string | number>
): Promise<void> {
  const tabla = await primeraTablaDeHoja(config, archivo, hoja);
  if (!tabla) {
    throw new ErrorSharePoint(
      `La hoja "${hoja}" no tiene ninguna Tabla de Excel definida — hace falta llamar a asegurarTablaEnHoja primero.`
    );
  }
  await graphFetch(
    config,
    `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/tables('${tabla.id}')/rows/add`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ values: [valores] }),
    }
  );
}

/**
 * Escribe una columna completa de fórmulas/valores de una sola vez (un solo rango,
 * una sola llamada) — para el cierre de mes, donde hay que ajustar la fórmula de
 * FORESCAST en decenas de filas a la vez sin hacer una escritura por fila.
 */
export async function escribirColumna(
  config: ConfiguracionSharePoint,
  archivo: ArchivoResuelto,
  hoja: string,
  columna: string,
  filaInicio: number,
  filaFin: number,
  valores: Array<string | number>
): Promise<void> {
  const direccion = `${columna}${filaInicio}:${columna}${filaFin}`;
  const ruta = `/drives/${archivo.driveId}/items/${archivo.itemId}/workbook/worksheets('${encodeURIComponent(
    hoja
  )}')/range(address='${direccion}')`;
  await graphFetch(config, ruta, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ formulas: valores.map((v) => [v]) }),
  });
}
