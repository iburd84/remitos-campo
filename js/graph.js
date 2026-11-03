// Acceso a SharePoint y Excel por Microsoft Graph, con la sesión del usuario.
import { CONFIG } from './config.js';
import { obtenerToken } from './auth.js';
import { planificarEscritura, filasDeRemito, columnasCliente, norm } from './excel.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const esperar = ms => new Promise(r => setTimeout(r, ms));

export async function g(ruta, { method = 'GET', body, headers = {}, reintentos = 3, scopes } = {}) {
  const h = { Authorization: 'Bearer ' + (await obtenerToken(scopes)), ...headers };
  let cuerpo = body;
  if (body !== undefined && !(body instanceof Blob)) {
    h['Content-Type'] = 'application/json';
    cuerpo = JSON.stringify(body);
  }
  const r = await fetch(ruta.startsWith('http') ? ruta : GRAPH + ruta, { method, headers: h, body: cuerpo });
  if ([429, 503, 504].includes(r.status) && reintentos > 0) {
    await esperar(1000 * (+r.headers.get('Retry-After') || 2));
    return g(ruta, { method, body, headers, reintentos: reintentos - 1, scopes });
  }
  if (!r.ok) {
    let msg = r.statusText, codigo;
    try { const j = await r.json(); msg = j.error?.message || msg; codigo = j.error?.code; } catch { /* sin cuerpo */ }
    const e = new Error(`SharePoint respondió ${r.status}: ${msg}`);
    e.status = r.status;
    e.codigo = codigo;
    throw e;
  }
  if (r.status === 204) return null;
  return (r.headers.get('content-type') || '').includes('json') ? r.json() : r;
}

async function todasLasPaginas(ruta, headers) {
  const out = [];
  for (let url = ruta; url;) {
    const r = await g(url, { headers });
    out.push(...r.value);
    url = r['@odata.nextLink'];
  }
  return out;
}

// ---------- Sitio, listas y biblioteca ----------

const cache = {};

async function sitioId() {
  if (!cache.sitio) cache.sitio = (await g(`/sites/${CONFIG.sitio.host}:${CONFIG.sitio.ruta}`)).id;
  return cache.sitio;
}

// Compara sin mayúsculas, acentos ni espacios: SharePoint puede guardar "Remitos Escaneados" como "RemitosEscaneados"
const clave = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[\s_-]+/g, '').toLowerCase();
const ultimoTramo = url => decodeURIComponent(String(url || '').replace(/\/(Forms\/AllItems\.aspx|AllItems\.aspx)?$/i, '').split('/').pop());

async function buscarLista(nombre) {
  const sid = await sitioId();
  const listas = await todasLasPaginas(`/sites/${sid}/lists?$select=id,name,displayName,webUrl`);
  cache.listasVistas = listas.map(l => l.displayName);
  const k = clave(nombre);
  return listas.find(l => [l.displayName, l.name, ultimoTramo(l.webUrl)].some(x => clave(x) === k)) || null;
}

async function listaId(nombre) {
  if (!cache[nombre]) {
    const l = await buscarLista(nombre);
    if (!l) {
      throw Object.assign(new Error(`Falta la lista "${nombre}". Un administrador debe correr la Configuración. ` +
        `(Listas del sitio: ${(cache.listasVistas || []).join(', ')})`), { codigo: 'sinConfigurar' });
    }
    cache[nombre] = l.id;
  }
  return cache[nombre];
}

async function driveId() {
  if (!cache.drive) {
    const sid = await sitioId();
    try {
      cache.drive = (await g(`/sites/${sid}/lists/${await listaId(CONFIG.biblioteca)}/drive?$select=id`)).id;
    } catch (e) {
      // Plan B: buscar la biblioteca entre las bibliotecas de documentos del sitio
      if (e.codigo !== 'sinConfigurar') throw e;
      const drives = await todasLasPaginas(`/sites/${sid}/drives?$select=id,name,webUrl`);
      const k = clave(CONFIG.biblioteca);
      const d = drives.find(x => clave(x.name) === k || clave(ultimoTramo(x.webUrl)) === k);
      if (!d) throw Object.assign(new Error(e.message + ` (Bibliotecas: ${drives.map(x => x.name).join(', ')})`), { codigo: 'sinConfigurar' });
      cache.drive = d.id;
    }
  }
  return cache.drive;
}

const rutaArchivo = p => p.split('/').map(encodeURIComponent).join('/');
const nombreSeguro = s => String(s).replace(/["*:<>?/\\|#%]/g, '-').replace(/\s+/g, ' ').trim();

// Orden de los PDF: <Cliente>/<Año>/<MM - Mes>/Remito <N°> - <Cliente> - <Establecimiento> - <AAAA-MM-DD>.pdf
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
const carpetaMes = fecha => { const [y, m] = fecha.split('-'); return `${y}/${m} - ${MESES[+m - 1]}`; };
const nombrePdf = (r, cliente) =>
  nombreSeguro(`Remito ${r.remito} - ${cliente} - ${r.establecimiento} - ${r.fechaRemito}`) + '.pdf';

// ---------- Campos de la lista <-> remito ----------

const fechaLista = iso => (iso ? iso + 'T12:00:00Z' : null);
function fechaLocal(v) {
  if (!v) return '';
  const d = new Date(v);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function aCampos(r) {
  return {
    Title: r.remito,
    Remito: r.remito,
    Cliente: r.cliente,
    Establecimiento: r.establecimiento,
    Zona: r.zona,
    FechaRemito: fechaLista(r.fechaRemito),
    FechaSolicitud: fechaLista(r.fechaSolicitud),
    Periodo: r.periodo || '',
    Tecnico: r.tecnico,
    TecnicoEmail: r.tecnicoEmail,
    // Con campos adicionales o fotos originales se guarda { items, extras, paginas }; si no, la lista sola (formato original)
    Items: JSON.stringify((r.extras && Object.keys(r.extras).length) || r.paginas?.length || r.gruposExtras?.length
      ? { items: r.items || [], extras: r.extras || {}, paginas: r.paginas || [], gruposExtras: r.gruposExtras || [] }
      : r.items || []),
    Kilometros: +r.km || 0,
    Observaciones: r.observaciones || '',
    Estado: r.estado,
    ComentarioRevision: r.comentario || '',
    Revisor: r.revisor || '',
    FechaRevision: r.fechaRevision || null,
    ArchivoId: r.archivoId || '',
    ArchivoUrl: r.archivoUrl || '',
    FilasExcel: r.filasExcel || '',
    IdLocal: r.idLocal || '',
  };
}

function deItem(it) {
  const f = it.fields;
  let items = [], extras = {}, paginas = [], gruposExtras = [];
  try {
    const j = JSON.parse(f.Items || '[]');
    if (Array.isArray(j)) items = j; else ({ items = [], extras = {}, paginas = [], gruposExtras = [] } = j);
  } catch { /* vacío */ }
  return {
    id: it.id,
    remito: f.Remito || f.Title,
    cliente: f.Cliente,
    establecimiento: f.Establecimiento,
    zona: f.Zona,
    fechaRemito: fechaLocal(f.FechaRemito),
    fechaSolicitud: fechaLocal(f.FechaSolicitud),
    periodo: f.Periodo || '',
    tecnico: f.Tecnico,
    tecnicoEmail: f.TecnicoEmail,
    items,
    extras,
    paginas,
    gruposExtras,
    km: f.Kilometros || 0,
    observaciones: f.Observaciones || '',
    estado: f.Estado,
    comentario: f.ComentarioRevision || '',
    revisor: f.Revisor || '',
    archivoId: f.ArchivoId || '',
    archivoUrl: f.ArchivoUrl || '',
    filasExcel: f.FilasExcel || '',
    idLocal: f.IdLocal || '',
    creado: it.createdDateTime,
  };
}

const PREFER = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };

async function itemsDonde(campo, valor) {
  const sid = await sitioId();
  const lid = await listaId(CONFIG.listaRemitos);
  const v = String(valor).replace(/'/g, "''");
  const items = await todasLasPaginas(
    `/sites/${sid}/lists/${lid}/items?$expand=fields&$top=200&$filter=fields/${campo} eq '${v}'`, PREFER);
  return items.map(deItem);
}

// ---------- Archivos ----------

// ruta relativa a la raíz de la biblioteca de trabajo, o a la carpeta `desde` = { drive, carpeta } si se indica
async function subirArchivo(ruta, blob, desde) {
  const base = desde
    ? `/drives/${desde.drive}/items/${desde.carpeta}:/${rutaArchivo(ruta)}:`
    : `/drives/${await driveId()}/root:/${rutaArchivo(ruta)}:`;
  if (blob.size < 3.5 * 1024 * 1024) {
    return g(`${base}/content?@microsoft.graph.conflictBehavior=replace`, {
      method: 'PUT', body: blob, headers: { 'Content-Type': blob.type || 'application/octet-stream' },
    });
  }
  const s = await g(`${base}/createUploadSession`, {
    method: 'POST', body: { item: { '@microsoft.graph.conflictBehavior': 'replace' } },
  });
  const PARTE = 5 * 320 * 1024;
  let r;
  for (let i = 0; i < blob.size; i += PARTE) {
    const trozo = blob.slice(i, i + PARTE);
    r = await fetch(s.uploadUrl, {
      method: 'PUT', body: trozo,
      headers: { 'Content-Range': `bytes ${i}-${i + trozo.size - 1}/${blob.size}` },
    });
    if (!r.ok && r.status !== 202) throw new Error('Falló la subida del PDF (' + r.status + ').');
  }
  return r.json();
}

async function descargar(driveIdArchivo, itemId) {
  const it = await g(`/drives/${driveIdArchivo}/items/${itemId}`);
  const r = await fetch(it['@microsoft.graph.downloadUrl']);
  if (!r.ok) throw new Error('No se pudo descargar el archivo.');
  return r.blob();
}

// ---------- Excel ----------

function wb(excel) { return `/drives/${excel.driveId}/items/${excel.itemId}/workbook`; }
const comillas = s => encodeURIComponent(String(s).replace(/'/g, "''"));

async function conSesion(excel, fn) {
  const s = await g(`${wb(excel)}/createSession`, { method: 'POST', body: { persistChanges: true } });
  const h = { 'workbook-session-id': s.id };
  try { return await fn(h); }
  finally { g(`${wb(excel)}/closeSession`, { method: 'POST', headers: h }).catch(() => {}); }
}

async function leerRango(excel, hoja, direccion, h) {
  return (await g(`${wb(excel)}/worksheets('${comillas(hoja)}')/range(address='${direccion}')?$select=values`, { headers: h })).values;
}

// ---------- API pública (misma forma que la de demo.js) ----------

export const api = {
  modo: 'sharepoint',

  async usuario() {
    const yo = await g('/me?$select=displayName,givenName,surname,mail,userPrincipalName');
    const email = (yo.mail || yo.userPrincipalName || '').toLowerCase();
    let fila = null;
    try {
      const sid = await sitioId();
      const lid = await listaId(CONFIG.listaUsuarios);
      const items = await todasLasPaginas(`/sites/${sid}/lists/${lid}/items?$expand=fields&$top=500`);
      fila = items.map(i => i.fields).find(f => (f.Title || '').toLowerCase() === email);
    } catch (e) {
      if (e.status === 403 || e.status === 404) {
        throw new Error(`Tu cuenta (${email}) no tiene acceso al sitio de SharePoint "${CONFIG.sitio.ruta.split('/').pop()}". ` +
          'Pedile al propietario del sitio que te agregue como miembro y volvé a entrar.');
      }
      if (e.codigo !== 'sinConfigurar') throw e;
      return { email, displayName: yo.displayName, nombre: yo.surname || yo.displayName, rol: 'Admin', sinConfigurar: true };
    }
    return {
      email,
      displayName: yo.displayName,
      nombre: fila?.NombrePlanilla || yo.surname || yo.displayName,
      rol: fila?.Rol || 'Tecnico',
    };
  },

  async leerClientes() {
    const did = await driveId();
    try {
      const it = await g(`/drives/${did}/root:/${rutaArchivo(CONFIG.archivoConfig)}`);
      const j = JSON.parse(await (await descargar(did, it.id)).text()); // descargar devuelve un Blob
      return j.clientes;
    } catch (e) {
      if (e.status === 404) throw Object.assign(new Error('Falta la configuración de clientes. Un administrador debe correr la Configuración.'), { codigo: 'sinConfigurar' });
      throw e;
    }
  },

  async guardarClientes(clientes) {
    const blob = new Blob([JSON.stringify({ clientes }, null, 2)], { type: 'application/json' });
    await subirArchivo(CONFIG.archivoConfig, blob);
  },

  async buscarRemito(nro) {
    return (await itemsDonde('Remito', nro))[0] || null;
  },

  // e = { idLocal, itemId?, datos, pdf? }
  async enviarRemito(e) {
    const d = e.datos;
    const existente = await this.buscarRemito(d.remito);
    if (existente && existente.idLocal !== e.idLocal && existente.id !== e.itemId) {
      throw Object.assign(new Error(`El remito ${d.remito} ya fue cargado por ${existente.tecnico}.`), { codigo: 'duplicado' });
    }
    const id = e.itemId || existente?.id;
    const datos = { ...d, idLocal: e.idLocal, estado: 'Pendiente', comentario: '' };
    if (e.pdf) {
      const carpeta = `${nombreSeguro(d.clienteNombre || d.cliente)}/${carpetaMes(d.fechaRemito)}`;
      const nombre = nombrePdf(d, d.clienteNombre || d.cliente);
      const arch = await subirArchivo(`${carpeta}/${nombre}`, e.pdf);
      datos.archivoId = arch.id;
      datos.archivoUrl = arch.webUrl;
      // Fotos originales con sus ajustes, para poder retocarlas en la revisión
      if (e.paginas?.length) {
        datos.paginas = [];
        for (const [i, p] of e.paginas.entries()) {
          const foto = await subirArchivo(`${carpeta}/_originales/${nombre.replace(/\.pdf$/, '')} - p${i + 1}.jpg`, p.blob);
          datos.paginas.push({ archivoId: foto.id, giro: p.giro || 0, recorte: p.recorte || null, mejorar: !!p.mejorar });
        }
      }
    }
    const sid = await sitioId();
    const lid = await listaId(CONFIG.listaRemitos);
    if (id) {
      const campos = aCampos(datos);
      if (!e.pdf) { delete campos.ArchivoId; delete campos.ArchivoUrl; }
      await g(`/sites/${sid}/lists/${lid}/items/${id}/fields`, { method: 'PATCH', body: campos });
    } else {
      await g(`/sites/${sid}/lists/${lid}/items`, { method: 'POST', body: { fields: aCampos(datos) } });
    }
  },

  async misRemitos(email) {
    return (await itemsDonde('TecnicoEmail', email))
      .sort((a, b) => (b.creado || '').localeCompare(a.creado || ''));
  },

  async pendientes() {
    const [a, b] = await Promise.all([itemsDonde('Estado', 'Pendiente'), itemsDonde('Estado', 'Pendiente Excel')]);
    return [...a, ...b].sort((x, y) => (x.creado || '').localeCompare(y.creado || ''));
  },

  async pdf(remito) {
    if (!remito.archivoId) return null;
    return descargar(await driveId(), remito.archivoId);
  },

  // Elimina el remito, su PDF y sus fotos originales. Todo va a la papelera de reciclaje del sitio.
  async eliminarRemito(remito) {
    const sid = await sitioId();
    const lid = await listaId(CONFIG.listaRemitos);
    const did = await driveId();
    const archivos = [remito.archivoId, ...(remito.paginas || []).map(p => p.archivoId)].filter(Boolean);
    for (const id of archivos) {
      await g(`/drives/${did}/items/${id}`, { method: 'DELETE' }).catch(e => { if (e.status !== 404) throw e; });
    }
    await g(`/sites/${sid}/lists/${lid}/items/${remito.id}`, { method: 'DELETE' });
  },

  // ---------- Retoque de fotos en la revisión ----------

  async fotoOriginal(pagina) {
    return descargar(await driveId(), pagina.archivoId);
  },

  // Sube una foto nueva agregada en la revisión; devuelve su id
  async subirFotoOriginal(remito, blob) {
    const cli = remito.clienteNombre || remito.cliente;
    const nombre = nombrePdf(remito, cli).replace(/\.pdf$/, '');
    const foto = await subirArchivo(`${nombreSeguro(cli)}/${carpetaMes(remito.fechaRemito)}/_originales/${nombre} - p${Date.now()}.jpg`, blob);
    return foto.id;
  },

  // Reemplaza el PDF del remito por el regenerado (o lo crea si no tenía)
  async reemplazarPdf(remito, pdf) {
    const did = await driveId();
    if (remito.archivoId) {
      const it = await g(`/drives/${did}/items/${remito.archivoId}/content`, {
        method: 'PUT', body: pdf, headers: { 'Content-Type': 'application/pdf' },
      });
      return { archivoId: it.id, archivoUrl: it.webUrl };
    }
    const cli = remito.clienteNombre || remito.cliente;
    const it = await subirArchivo(`${nombreSeguro(cli)}/${carpetaMes(remito.fechaRemito)}/${nombrePdf(remito, cli)}`, pdf);
    return { archivoId: it.id, archivoUrl: it.webUrl };
  },

  async actualizar(remito) {
    const sid = await sitioId();
    const lid = await listaId(CONFIG.listaRemitos);
    const campos = aCampos(remito);
    delete campos.IdLocal;
    await g(`/sites/${sid}/lists/${lid}/items/${remito.id}/fields`, { method: 'PATCH', body: campos });
  },

  // Escribe el remito en la planilla del cliente. Devuelve "desde-hasta".
  async escribirExcel(cliente, remito) {
    const ex = cliente.excel;
    const filas = filasDeRemito(remito, cliente.items);
    return conSesion(ex, async h => {
      const t = `${wb(ex)}/tables('${comillas(ex.tabla)}')`;
      const [enc, cuerpo] = await Promise.all([
        g(`${t}/headerRowRange?$select=values`, { headers: h }),
        g(`${t}/dataBodyRange?$select=address,formulas,numberFormat`, { headers: h }),
      ]);
      const plan = planificarEscritura({
        encabezados: enc.values[0], cuerpo: cuerpo.formulas, formatos: cuerpo.numberFormat,
        direccion: cuerpo.address, columnas: columnasCliente(cliente), filas,
      });
      if (plan.error) throw Object.assign(new Error(plan.error), { codigo: plan.duplicado ? 'duplicado' : 'planilla' });
      for (const { address, ...cambios } of plan.escrituras) {
        await g(`${wb(ex)}/worksheets('${comillas(plan.hoja)}')/range(address='${address}')`, {
          method: 'PATCH', body: cambios, headers: h,
        });
      }
      if (plan.agregar.length) {
        await g(`${t}/rows/add`, { method: 'POST', body: { values: plan.agregar }, headers: h });
      }
      return `${plan.desde}-${plan.hasta}`;
    });
  },

  // Copia el PDF aprobado al lado del Excel del cliente, con los datos ya revisados en el nombre.
  // Devuelve el link del archivo copiado.
  async archivarPdf(cliente, remito) {
    if (!remito.archivoId) throw new Error('El remito no tiene foto.');
    const ex = cliente.excel;
    const blob = await descargar(await driveId(), remito.archivoId);
    const padre = (await g(`/drives/${ex.driveId}/items/${ex.itemId}?$select=parentReference`)).parentReference.id;
    const ruta = `Remitos ${nombreSeguro(cliente.nombre)}/${carpetaMes(remito.fechaRemito)}/${nombrePdf(remito, cliente.nombre)}`;
    const it = await subirArchivo(ruta, new Blob([blob], { type: 'application/pdf' }), { drive: ex.driveId, carpeta: padre });
    return it.webUrl;
  },

  // ---------- Configuración inicial ----------

  async crearEstructura(alAvanzar) {
    const sid = await sitioId();
    const texto = (name, extra = {}) => ({ name, text: {}, ...extra });
    const fecha = name => ({ name, dateTime: { format: 'dateOnly' } });
    const largo = name => ({ name, text: { allowMultipleLines: true } });
    const definiciones = [
      {
        displayName: CONFIG.listaRemitos, list: { template: 'genericList' },
        columns: [
          texto('Remito', { indexed: true }), texto('Cliente'), texto('Establecimiento'), texto('Zona'),
          fecha('FechaRemito'), fecha('FechaSolicitud'), texto('Periodo'),
          texto('Tecnico'), texto('TecnicoEmail', { indexed: true }), largo('Items'),
          { name: 'Kilometros', number: {} }, largo('Observaciones'),
          { name: 'Estado', indexed: true, choice: { choices: ['Pendiente', 'Pendiente Excel', 'Aprobado', 'Rechazado'] } },
          largo('ComentarioRevision'), texto('Revisor'), { name: 'FechaRevision', dateTime: {} },
          texto('ArchivoId'), texto('ArchivoUrl'), texto('FilasExcel'), texto('IdLocal'),
        ],
      },
      {
        displayName: CONFIG.listaUsuarios, list: { template: 'genericList' },
        columns: [texto('NombrePlanilla'), { name: 'Rol', choice: { choices: ['Tecnico', 'Revisor', 'Admin'] } }],
      },
      { displayName: CONFIG.biblioteca, list: { template: 'documentLibrary' } },
    ];
    for (const def of definiciones) {
      if (await buscarLista(def.displayName)) { alAvanzar?.(`"${def.displayName}" ya existía.`); continue; }
      // Crear listas requiere Sites.Manage.All (Sites.ReadWrite.All solo trabaja dentro de ellas)
      await g(`/sites/${sid}/lists`, { method: 'POST', body: def, scopes: CONFIG.scopesConfiguracion });
      alAvanzar?.(`Creada "${def.displayName}".`);
    }
    Object.keys(cache).forEach(k => k !== 'sitio' && delete cache[k]);
  },

  async agregarUsuario(email, nombrePlanilla, rol) {
    const sid = await sitioId();
    const lid = await listaId(CONFIG.listaUsuarios);
    await g(`/sites/${sid}/lists/${lid}/items`, {
      method: 'POST', body: { fields: { Title: email.toLowerCase(), NombrePlanilla: nombrePlanilla, Rol: rol } },
    });
  },

  async urlLista(nombre) { return (await buscarLista(nombre))?.webUrl; },

  // Resuelve el link "Compartir/Copiar vínculo" del Excel
  async resolverExcel(link) {
    const b64 = btoa(unescape(encodeURIComponent(link.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const it = await g(`/shares/u!${b64}/driveItem?$select=id,name,webUrl,parentReference`)
      .catch(e => { throw new Error('No se pudo abrir el vínculo del Excel. Copialo de nuevo desde SharePoint (··· → Copiar vínculo). ' + e.message); });
    return { driveId: it.parentReference.driveId, itemId: it.id, nombre: it.name, webUrl: it.webUrl };
  },

  // Cuenta cuántas veces se cargó cada ítem en las últimas filas de la planilla (para "Más usados")
  async contarItems(excel, tabla, encabezadoItem, ultimas = 500) {
    return conSesion(excel, async h => {
      const t = `${wb(excel)}/tables('${comillas(tabla)}')`;
      const enc = (await g(`${t}/headerRowRange?$select=values`, { headers: h })).values[0];
      const i = enc.map(norm).indexOf(norm(encabezadoItem));
      if (i < 0) return {};
      const v = (await g(`${t}/columns/itemAt(index=${i})/dataBodyRange?$select=values`, { headers: h })).values
        .map(f => f[0]).filter(x => x !== '' && x !== null);
      const cuenta = {};
      for (const cod of v.slice(-ultimas)) cuenta[cod] = (cuenta[cod] || 0) + 1;
      return cuenta;
    });
  },

  // Lee una tabla auxiliar completa (para las opciones de los campos adicionales)
  async leerTabla(excel, tabla) {
    return conSesion(excel, async h => {
      const v = (await g(`${wb(excel)}/tables('${comillas(tabla)}')/range?$select=values`, { headers: h })
        .catch(e => { throw new Error(`No se encontró la tabla "${tabla}" de un campo adicional en "${excel.nombre}". ${e.status === 404 ? '' : e.message}`); })).values;
      return { encabezados: v[0], filas: v.slice(1) };
    });
  },

  // Lee ítems (tabla de tarifas) y establecimientos (hoja y rango) desde el mismo Excel
  async leerCatalogo(excel, { tablaTarifas, hojaEstab, rangoEstab }) {
    return conSesion(excel, async h => {
      const enExcel = `en "${excel.nombre}"`;
      const tar = (await g(`${wb(excel)}/tables('${comillas(tablaTarifas)}')/range?$select=values`, { headers: h })
        .catch(e => { throw new Error(`No se encontró la tabla de ítems "${tablaTarifas}" ${enExcel}. ${e.status === 404 ? '' : e.message}`); })).values;
      // Establecimientos opcionales: sin hoja o rango, el técnico escribe el lugar a mano
      const est = hojaEstab && rangoEstab
        ? await leerRango(excel, hojaEstab, rangoEstab, h).catch(e => {
          throw new Error(`No se encontró la hoja "${hojaEstab}" ${enExcel}. Si este cliente no tiene lista de establecimientos, dejá la hoja vacía. ${e.status === 404 ? '' : e.message}`);
        })
        : [];
      return {
        items: tar.slice(1).filter(f => f[0] !== '' && f[1] !== '').map(f => ({ cod: f[0], servicio: String(f[1]).trim(), planilla: true })),
        // Se corta en la primera fila vacía: debajo suele haber otra tabla de la misma hoja
        establecimientos: est.slice(0, (est.findIndex(f => !String(f[1]).trim()) + 1 || est.length + 1) - 1)
          .map(f => ({ zona: String(f[0]).trim(), nombre: String(f[1]).trim() })),
      };
    });
  },
};
