// Arranque, sesión, navegación y utilidades de interfaz compartidas.
import { CONFIG, MODO_DEMO } from './config.js';
import { cola, procesarCola } from './cola.js';
import { masUsados, UNIDADES, camposDeGrupo, valoresDeGrupo } from './excel.js';

export const estado = { api: null, usuario: null, clientes: null };

// ---------- Utilidades ----------

export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const hoy = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
export const fechaCorta = iso => (iso ? iso.split('-').reverse().join('/') : '—');
export const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

export function recordar(clave, valor) {
  try {
    if (valor === undefined) return JSON.parse(localStorage.getItem('rc-' + clave));
    localStorage.setItem('rc-' + clave, JSON.stringify(valor));
  } catch { return null; }
  return valor;
}

export function aviso(msg, tipo = '') {
  const el = document.createElement('div');
  el.className = 'aviso ' + tipo;
  el.textContent = msg;
  $('#avisos').append(el);
  setTimeout(() => el.remove(), tipo === 'error' ? 7000 : 4000);
}

export function claseEstado(e) {
  return { Pendiente: 'pendiente', 'Pendiente Excel': 'pendiente', Aprobado: 'aprobado', Rechazado: 'rechazado' }[e] || 'cola';
}

// N° de remito: "5780" -> "0001-00005780"
export function normalizarRemito(pv, num) {
  const p = String(pv || '').replace(/\D/g, '');
  const n = String(num || '').replace(/\D/g, '');
  if (!p || !n) return '';
  return p.padStart(4, '0') + '-' + n.padStart(8, '0');
}

// Lista de servicios con botones +/-; cant = { cod: número }
function filaServicio(it, cant) {
  const v = +cant[it.cod] || 0;
  const km = it.unidad === 'km' || String(it.cod).toUpperCase() === 'KM';
  const detalle = km
    ? (it.planilla === false ? 'Kilómetros recorridos (no va a la planilla)' : `Ítem ${esc(it.cod)} · kilómetros`)
    : 'Ítem ' + esc(it.cod);
  return `<div class="servicio ${v > 0 ? 'marcado' : ''}" data-cod="${esc(it.cod)}">
    <div class="nombre">${esc(nombreServicio(it))}<small>${detalle}</small></div>
    <div class="stepper">
      ${km ? '' : '<button type="button" data-paso="-1" aria-label="Restar">−</button>'}
      <input type="number" inputmode="numeric" min="0" class="${km ? 'km' : ''}" value="${v || ''}" placeholder="0" aria-label="Cantidad ${esc(it.servicio)}">
      ${km ? '' : '<button type="button" data-paso="1" aria-label="Sumar">+</button>'}
    </div></div>`;
}

const lista = (items, cant) => `<div class="servicios">${items.map(it => filaServicio(it, cant)).join('')}</div>`;
const desplegable = (titulo, items, cant) => {
  const abierto = items.some(it => +cant[it.cod] > 0); // si tiene algo cargado, se muestra abierto
  return `<details class="otros-servicios" ${abierto ? 'open' : ''}><summary>${esc(titulo)} (${items.length})</summary>${lista(items, cant)}</details>`;
};

// Ítems por unidad de negocio: comunes a la vista; Mediciones laborales y Seguridad eléctrica en desplegables
function porUnidad(items, cant) {
  const comunes = items.filter(it => it.negocio === 'comun');
  return (comunes.length ? `<p class="etiqueta">${UNIDADES.comun}</p>${lista(comunes, cant)}` : '') +
    ['ml', 'se'].map(k => {
      const g = items.filter(it => (it.negocio || 'ml') === k);
      return g.length ? desplegable(UNIDADES[k], g, cant) : '';
    }).join('');
}

// ---------- Grupos de un remito (ej. una línea de Halliburton cada uno) ----------

export function chipsGrupo(defs, valores) {
  const v = valoresDeGrupo(defs, valores);
  return camposDeGrupo(defs).filter(d => d.quien === 'auto' && v[d.campo] !== '')
    .map(d => `${esc(d.etiqueta)} <span class="chip">${esc(v[d.campo])}</span>`).join(' ');
}

// Cabecera del grupo g: sus campos para elegir (ej. Línea) y lo que se calcula de ellos (ej. Centro de costo)
export function htmlCabeceraGrupo(defs, valores, g, total) {
  const elegibles = camposDeGrupo(defs).filter(d => d.quien !== 'auto');
  const titulo = elegibles[0]?.etiqueta || 'Grupo';
  return `<div class="grupo-cab">
    ${total > 1 ? `<div class="fila" style="justify-content:space-between"><b>${esc(titulo)} ${g + 1}</b>
      ${g > 0 ? `<button type="button" class="peligro" data-quitar-grupo="${g}">Quitar</button>` : ''}</div>` : ''}
    ${elegibles.map(d => {
      const v = valores?.[d.campo] || '';
      const ops = !v || (d.opciones || []).includes(v) ? d.opciones || [] : [v, ...(d.opciones || [])];
      return `<label>${esc(d.etiqueta)}${d.obligatorio ? '' : ' (opcional)'}</label>
        <select data-g="${g}" data-gextra="${esc(d.campo)}"><option value="">Elegir…</option>
        ${ops.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
    }).join('')}
    <p class="ayuda" data-autos-g="${g}">${chipsGrupo(defs, valores)}</p>
  </div>`;
}

// Formulario del técnico: "Más usados" arriba; el resto por unidad de negocio o en un desplegable
export function htmlServiciosTecnico(cli, cant, usoLocal) {
  const top = masUsados(cli.items, usoLocal);
  if (!top.length && !cli.separarUnidades) return lista(cli.items, cant);
  const enTop = new Set(top.map(it => String(it.cod)));
  const resto = cli.items.filter(it => !enTop.has(String(it.cod)));
  return (top.length ? `<p class="etiqueta">Más usados</p>${lista(top, cant)}` : '') +
    (cli.separarUnidades ? porUnidad(resto, cant) : desplegable('Todos los demás servicios', resto, cant));
}

// separarCargados (revisión): primero lo que cargó el técnico; el resto en un desplegable para sumar otro
export function htmlServicios(items, cant, { separarCargados = false, unidades = false } = {}) {
  if (!separarCargados) return lista(items, cant);
  const cargados = items.filter(it => +cant[it.cod] > 0);
  const resto = items.filter(it => !(+cant[it.cod] > 0));
  return `${cargados.length ? lista(cargados, cant) : '<p class="ayuda">El técnico no cargó servicios.</p>'}
    ${resto.length ? `<details class="otros-servicios"><summary>Agregar otro servicio (${resto.length})</summary>
      ${unidades ? porUnidad(resto, cant) : lista(resto, cant)}</details>` : ''}`;
}

export function nombreServicio(it) {
  const s = String(it.servicio || '').toLowerCase();
  return s.charAt(0).toUpperCase() + s.slice(1);
}

export function conectarServicios(raiz, cant, alCambiar) {
  raiz.addEventListener('click', e => {
    const b = e.target.closest('button[data-paso]');
    if (!b) return;
    const fila = b.closest('.servicio');
    const inp = $('input', fila);
    inp.value = Math.max(0, (+inp.value || 0) + +b.dataset.paso) || '';
    inp.dispatchEvent(new Event('input', { bubbles: true }));
  });
  raiz.addEventListener('input', e => {
    const fila = e.target.closest('.servicio');
    if (!fila) return;
    const v = Math.max(0, Math.floor(+e.target.value || 0));
    cant[fila.dataset.cod] = v;
    fila.classList.toggle('marcado', v > 0);
    alCambiar?.();
  });
}

// ---------- Navegación ----------

const RUTAS = {
  nuevo: { titulo: 'Nuevo remito', roles: ['Tecnico', 'Revisor', 'Admin'], cargar: () => import('./tecnico.js').then(m => m.vistaNuevo) },
  mis: { titulo: 'Mis remitos', roles: ['Tecnico', 'Revisor', 'Admin'], cargar: () => import('./tecnico.js').then(m => m.vistaMisRemitos) },
  revision: { titulo: 'Revisión', roles: ['Revisor', 'Admin'], cargar: () => import('./revisor.js').then(m => m.vistaRevision) },
  config: { titulo: 'Configuración', roles: ['Admin'], cargar: () => import('./setup.js').then(m => m.vistaConfiguracion) },
};

export function ir(ruta) {
  if (location.hash === '#' + ruta) mostrar(); else location.hash = ruta;
}

async function pintarPestanas() {
  const nav = $('#pestanas');
  const pendientes = (await cola.todos().catch(() => [])).length;
  const actual = (location.hash.slice(1).split('/')[0]) || 'nuevo';
  nav.innerHTML = Object.entries(RUTAS)
    .filter(([, r]) => r.roles.includes(estado.usuario.rol))
    .map(([k, r]) => `<a href="#${k}" class="${k === actual ? 'activa' : ''}">${r.titulo}${k === 'mis' && pendientes ? `<span class="contador" title="Sin enviar">${pendientes}</span>` : ''}</a>`)
    .join('');
  nav.hidden = false;
}

export async function mostrar() {
  if (!estado.usuario) return;
  const [ruta, ...resto] = location.hash.slice(1).split('/');
  let r = RUTAS[ruta];
  if (!r || !r.roles.includes(estado.usuario.rol)) r = RUTAS.nuevo;
  await pintarPestanas();
  const vista = $('#vista');
  vista.className = 'vista';
  try {
    const fn = await r.cargar();
    await fn(vista, resto.join('/'));
  } catch (e) {
    console.error(e);
    vista.innerHTML = `<div class="mensaje error">${esc(e.message)}</div>`;
  }
}

export const refrescarPestanas = () => estado.usuario && pintarPestanas();

// ---------- Arranque ----------

async function enviarPendientes() {
  await procesarCola(estado.api, refrescarPestanas);
}

async function iniciar() {
  $('#version').textContent = 'v' + CONFIG.version;
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' })
      .then(reg => reg.update())
      .catch(e => console.warn('Sin modo offline:', e));
    // Cuando se instala una versión nueva, recargar una vez para usarla
    const habiaControlador = !!navigator.serviceWorker.controller;
    let recargado = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (habiaControlador && !recargado) { recargado = true; location.reload(); }
    });
  }

  if (MODO_DEMO) {
    estado.api = (await import('./demo.js')).api;
    $('#banner').innerHTML = '<div class="banner">Modo demostración: los datos quedan en este navegador y la planilla es simulada.</div>';
  } else {
    const auth = await import('./auth.js');
    let cuenta;
    try { cuenta = await auth.iniciarAuth(); } catch (e) { $('#vista').innerHTML = `<div class="mensaje error">No se pudo iniciar el login: ${esc(e.message)}</div>`; return; }
    if (!cuenta) {
      $('#vista').innerHTML = `<div class="tarjeta login">
        <h2>Ingresar</h2><p>Usá tu cuenta corporativa de Alonger (Microsoft 365).</p>
        <button class="primario grande ancho" id="btn-ingresar">Ingresar con Microsoft</button></div>`;
      $('#btn-ingresar').onclick = () => auth.ingresar();
      return;
    }
    estado.api = (await import('./graph.js')).api;
    estado.salir = auth.salir;
  }

  // Usuario y clientes: si no hay señal, se usa lo último guardado en el celular
  try {
    estado.usuario = recordar('usuario', await estado.api.usuario());
  } catch (e) {
    estado.usuario = recordar('usuario');
    if (!estado.usuario) {
      $('#vista').innerHTML = `<div class="mensaje error">No se pudo identificar al usuario: ${esc(e.message)}</div>` +
        (estado.salir ? '<div class="acciones"><button id="btn-salir-error">Salir / cambiar de cuenta</button></div>' : '');
      $('#btn-salir-error')?.addEventListener('click', () => estado.salir());
      return;
    }
  }
  try {
    estado.clientes = recordar('clientes', await estado.api.leerClientes());
  } catch (e) {
    estado.clientes = recordar('clientes');
    if (!estado.clientes) estado.errorClientes = e.message;
  }

  $('#usuario').innerHTML = `${esc(estado.usuario.displayName)}<br><small>${esc(estado.usuario.rol)}</small>` +
    (estado.salir ? ` · <button id="btn-salir">Salir</button>` : '');
  $('#btn-salir')?.addEventListener('click', () => estado.salir());

  if (!estado.clientes && estado.usuario.rol === 'Admin' && !location.hash.startsWith('#config')) location.hash = 'config';
  window.addEventListener('hashchange', mostrar);
  window.addEventListener('online', enviarPendientes);
  setInterval(enviarPendientes, 60000);
  await mostrar();
  enviarPendientes();
}

iniciar();
