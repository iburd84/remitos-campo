// Pantalla de revisión (oficina): foto del remito al lado de los datos, corregir, aprobar o rechazar.
// Lo que corrige el revisor tiene prioridad sobre lo que cargó el técnico.
import {
  estado, $, esc, fechaCorta, recordar, aviso, ir, claseEstado, normalizarRemito,
  htmlServicios, conectarServicios, htmlCabeceraGrupo, chipsGrupo,
} from './app.js';
import { completarAutomaticos, camposDeGrupo, usaGrupos, valoresDeGrupo, gruposDeRemito } from './excel.js';
import { editarFoto, miniatura } from './fotos.js';
import { prepararFoto, armarPdf } from './pdf.js';

let urlPdf = null;

export async function vistaRevision(vista, id) {
  if (id) return detalle(vista, id);
  vista.className = 'vista ancha';
  vista.innerHTML = '<h1 class="titulo">Remitos para revisar</h1><div id="lista"><p class="vacio">Cargando…</p></div>';
  try {
    const lista = await estado.api.pendientes();
    estado.pendientes = lista;
    $('#lista', vista).innerHTML = lista.length ? `<div class="lista-remitos">${lista.map(r => `
      <a class="remito-fila boton" href="#revision/${esc(r.id)}" style="justify-content:flex-start">
        <div class="datos"><div class="nro">${esc(r.remito)} · ${esc(estado.clientes?.[r.cliente]?.nombre || r.cliente)}</div>
          <div class="sub">${esc(r.establecimiento)} · ${fechaCorta(r.fechaRemito)} · ${esc(r.tecnico)}</div>
          ${r.estado === 'Pendiente Excel' && r.comentario ? `<div class="sub" style="color:var(--error)">${esc(r.comentario)}</div>` : ''}</div>
        <span class="estado ${claseEstado(r.estado)}">${esc(r.estado)}</span>
      </a>`).join('')}</div>` : '<p class="vacio">No hay remitos pendientes. ✔</p>';
  } catch (e) {
    $('#lista', vista).innerHTML = `<div class="mensaje error">${esc(e.message)}</div>`;
  }
}

async function detalle(vista, id) {
  vista.className = 'vista ancha';
  let r = (estado.pendientes || []).find(x => x.id === id);
  if (!r) r = (await estado.api.pendientes()).find(x => x.id === id);
  if (!r) { vista.innerHTML = '<div class="mensaje alerta">Ese remito ya no está pendiente.</div><a href="#revision" class="boton">Volver</a>'; return; }
  const cli = estado.clientes?.[r.cliente];
  if (!cli) { vista.innerHTML = `<div class="mensaje error">El cliente "${esc(r.cliente)}" no está configurado.</div>`; return; }

  // Valores por defecto para lo que completa la oficina
  r = { ...r };
  if (!r.fechaSolicitud) r.fechaSolicitud = recordar('ultimaFechaSolicitud') || '';
  if (!r.periodo) r.periodo = recordar('ultimoPeriodo') || r.fechaRemito.slice(0, 7);
  const [pv, numero] = r.remito.split('-');
  // Servicios por grupo (ej. una línea de Halliburton cada uno); campos de grupo y campos por remito
  const conGrupos = usaGrupos(cli);
  const grupos = gruposDeRemito(r, cli.extras);
  const deGrupo = new Set(camposDeGrupo(cli.extras).map(d => d.campo));
  const extrasRemito = (cli.extras || []).filter(d => !deGrupo.has(d.campo));
  const etiquetaGrupo = (camposDeGrupo(cli.extras).find(d => d.quien !== 'auto')?.etiqueta || 'grupo').toLowerCase();

  vista.innerHTML = `
    <p><a href="#revision">← Volver a la lista</a></p>
    <div class="revision">
      <div class="visor" id="visor"><p class="vacio">Cargando foto…</p></div>
      <div>
        ${r.estado === 'Pendiente Excel' ? `<div class="mensaje error"><b>No se pudo pasar al Excel:</b> ${esc(r.comentario)}</div>` : ''}
        <section class="tarjeta">
          <h2>${esc(cli.nombre)} — remito ${esc(r.remito)}</h2>
          <p class="ayuda" style="margin-top:-8px">Cargado por ${esc(r.tecnico)}${r.creado ? ' el ' + new Date(r.creado).toLocaleString('es-AR') : ''}</p>
          <label>N° de remito</label>
          <div class="fila nro-remito">
            <input id="f-pv" class="pv" inputmode="numeric" value="${esc(pv)}"><span class="guion">–</span>
            <input id="f-numero" inputmode="numeric" value="${esc(numero)}">
          </div>
          <div class="dos-col">
            <div><label for="f-fecha">Fecha del remito</label><input id="f-fecha" type="date" value="${esc(r.fechaRemito)}"></div>
            <div><label for="f-tecnico">Responsable técnico</label><input id="f-tecnico" value="${esc(r.tecnico)}"></div>
            <div><label for="f-solicitud">Fecha de solicitud</label><input id="f-solicitud" type="date" value="${esc(r.fechaSolicitud)}"></div>
            <div><label for="f-periodo">Período</label><input id="f-periodo" type="month" value="${esc(r.periodo)}"></div>
          </div>
          <label for="f-estab">Establecimiento / Pozo</label>
          ${cli.establecimientos?.length ? `<select id="f-estab">${cli.establecimientos.map(e => `<option ${e.nombre === r.establecimiento ? 'selected' : ''}>${esc(e.nombre)}</option>`).join('')}
            ${cli.establecimientos.some(e => e.nombre === r.establecimiento) ? '' : `<option selected>${esc(r.establecimiento)}</option>`}</select>`
            : `<input id="f-estab" value="${esc(r.establecimiento)}">`}
          <p class="ayuda" id="zona">${r.zona ? `Zona <span class="chip">${esc(r.zona)}</span>` : ''}</p>
          ${extrasRemito.filter(d => d.quien !== 'auto').map(d => {
            const v = r.extras?.[d.campo] || (d.opciones?.length === 1 ? d.opciones[0] : '');
            const ops = (d.opciones || []).includes(v) || !v ? d.opciones || [] : [v, ...(d.opciones || [])];
            return `<label for="x-${esc(d.campo)}">${esc(d.etiqueta)}${d.quien === 'revisor' ? ' (lo completa la oficina)' : ''}</label>
              <select id="x-${esc(d.campo)}" data-extra="${esc(d.campo)}"><option value="">Elegir…</option>
              ${ops.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
          }).join('')}
          <p class="ayuda" id="autos"></p>
        </section>
        <section class="tarjeta">
          <h2>Foto del remito</h2>
          <div class="fotos fotos-revision" id="fotos-rev"></div>
          <p class="ayuda" id="fotos-ayuda"></p>
          <div class="acciones">
            <label class="boton subir-foto">Agregar foto<input type="file" accept="image/*" id="f-agregar-foto"></label>
            <button id="btn-guardar-foto" disabled>Guardar foto</button>
          </div>
        </section>
        <section class="tarjeta">
          <h2>Servicios</h2>
          <p class="ayuda" style="margin:-4px 0 8px">Resaltado: lo que cargó el técnico.</p>
          <div id="servicios"></div>
          ${conGrupos ? `<div class="acciones"><button type="button" id="btn-grupo">+ Agregar otra ${esc(etiquetaGrupo)}</button></div>` : ''}
          ${r.observaciones ? `<label>Observaciones del técnico</label><p style="margin:0">${esc(r.observaciones)}</p>` : ''}
        </section>
        <section class="tarjeta">
          <h2>Resultado</h2>
          <label for="f-comentario">Comentario (obligatorio para rechazar)</label>
          <textarea id="f-comentario" rows="2"></textarea>
          <div id="msg"></div>
          <div class="acciones">
            <button class="primario grande" id="btn-aprobar">Aprobar y pasar al Excel</button>
          </div>
          <div class="acciones">
            <button id="btn-guardar">Guardar cambios</button>
            <button class="peligro" id="btn-rechazar">Rechazar</button>
          </div>
          <div class="acciones">
            <button class="peligro" id="btn-eliminar">Eliminar remito</button>
          </div>
          <p class="ayuda">Excel: ${esc(cli.excel?.nombre || '—')}</p>
        </section>
      </div>
    </div>`;

  function pintarServicios() {
    $('#servicios', vista).innerHTML = grupos.map((gr, g) => `<div class="${conGrupos ? 'grupo' : ''}" data-grupo="${g}">
      ${conGrupos ? htmlCabeceraGrupo(cli.extras, gr.extras, g, grupos.length) : ''}
      <div class="servicios-g">${htmlServicios(cli.items, gr.cant, { separarCargados: true, unidades: !!cli.separarUnidades })}</div>
    </div>`).join('');
    vista.querySelectorAll('#servicios [data-grupo]').forEach(div =>
      conectarServicios($('.servicios-g', div), grupos[+div.dataset.grupo].cant));
    vista.querySelectorAll('#servicios [data-gextra]').forEach(sel => sel.addEventListener('input', () => {
      const gr = grupos[+sel.dataset.g];
      gr.extras[sel.dataset.gextra] = sel.value;
      vista.querySelector(`[data-autos-g="${sel.dataset.g}"]`).innerHTML = chipsGrupo(cli.extras, gr.extras);
    }));
    vista.querySelectorAll('#servicios [data-quitar-grupo]').forEach(btn => btn.addEventListener('click', () => {
      const g = +btn.dataset.quitarGrupo;
      if (Object.values(grupos[g].cant).some(v => +v > 0) && !confirm('Este grupo tiene servicios. ¿Quitarlo igual?')) return;
      grupos.splice(g, 1);
      pintarServicios();
    }));
  }
  pintarServicios();
  $('#btn-grupo', vista)?.addEventListener('click', () => { grupos.push({ extras: {}, cant: {} }); pintarServicios(); });

  $('#f-estab', vista).addEventListener('input', e => {
    const est = (cli.establecimientos || []).find(x => x.nombre === e.target.value);
    r.zona = est?.zona || r.zona;
    $('#zona', vista).innerHTML = r.zona ? `Zona <span class="chip">${esc(r.zona)}</span>` : '';
  });

  // Campos por remito (ej. Sector): lo elegido más lo que se calcula de eso; los de grupo van aparte
  const leerExtras = () => {
    const v = completarAutomaticos(cli.extras, Object.fromEntries(
      [...vista.querySelectorAll('[data-extra]')].map(s => [s.dataset.extra, s.value])));
    return Object.fromEntries(extrasRemito.map(d => [d.campo, v[d.campo] ?? '']));
  };
  const pintarAutos = () => {
    const v = leerExtras();
    $('#autos', vista).innerHTML = extrasRemito.filter(d => d.quien === 'auto' && v[d.campo] !== '')
      .map(d => `${esc(d.etiqueta)} <span class="chip">${esc(v[d.campo])}</span>`).join(' ');
  };
  vista.querySelectorAll('[data-extra]').forEach(s => s.addEventListener('input', pintarAutos));
  pintarAutos();

  // Visor del PDF
  const cargarVisor = blobDirecto => {
    if (urlPdf) URL.revokeObjectURL(urlPdf);
    (blobDirecto ? Promise.resolve(blobDirecto) : estado.api.pdf(r)).then(blob => {
      const visor = $('#visor', vista);
      if (!visor) return;
      if (!blob) { visor.innerHTML = '<p class="vacio">El remito no tiene foto.</p>'; return; }
      urlPdf = URL.createObjectURL(blob);
      visor.innerHTML = `<iframe src="${urlPdf}#view=FitH" title="Remito escaneado"></iframe>`;
    }).catch(e => { $('#visor', vista).innerHTML = `<div class="mensaje error">${esc(e.message)}</div>`; });
  };
  cargarVisor();

  // Fotos originales: retocar (recorte, giro, mejora) o agregar, y regenerar el PDF
  let fotos = (r.paginas || []).map(m => ({ ...m, blob: null }));
  let fotosCambiadas = false;
  const ayudaFotos = () => {
    $('#fotos-ayuda', vista).textContent = fotos.length
      ? 'Tocá una foto para recortarla, girarla o mejorarla.' + (fotosCambiadas ? ' Hay cambios sin guardar.' : '')
      : (r.archivoId ? 'Este remito se cargó antes de guardar las fotos originales: no se puede retocar, pero podés agregar una foto nueva para reemplazar el PDF.' : '');
    $('#btn-guardar-foto', vista).disabled = !fotosCambiadas || !fotos.length;
  };
  async function pintarFotos() {
    const cont = $('#fotos-rev', vista);
    cont.innerHTML = '';
    for (const [i, f] of fotos.entries()) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'foto' + (f.cambiada ? ' cambiada' : '');
      btn.innerHTML = `<span class="num">${i + 1}</span>`;
      btn.onclick = async () => {
        if (!f.blob) return;
        const original = { blob: f.blob, giro: f.giro, recorte: f.recorte, mejorar: f.mejorar };
        const editada = await editarFoto(original);
        if (editada === original) return; // canceló
        if (editada === null) fotos.splice(i, 1);
        else Object.assign(f, editada, { cambiada: true });
        fotosCambiadas = true;
        pintarFotos();
      };
      cont.append(btn);
      (async () => {
        try {
          if (!f.blob) f.blob = await estado.api.fotoOriginal(f);
          const im = new Image();
          im.src = await miniatura(f);
          im.alt = 'Foto ' + (i + 1);
          btn.prepend(im);
        } catch { btn.insertAdjacentHTML('beforeend', '<span class="ayuda">No se pudo cargar</span>'); }
      })();
    }
    ayudaFotos();
  }
  pintarFotos();

  $('#f-agregar-foto', vista).addEventListener('change', async e => {
    const archivo = e.target.files[0];
    e.target.value = '';
    if (!archivo) return;
    const blob = await prepararFoto(archivo);
    const editada = await editarFoto({ blob, giro: 0, recorte: { x: 0, y: 0, w: 1, h: 1 }, mejorar: true });
    if (!editada) return;
    fotos.push({ ...editada, nueva: true, cambiada: true });
    fotosCambiadas = true;
    pintarFotos();
  });

  // Sube las fotos nuevas, regenera el PDF y lo reemplaza. Devuelve los datos de archivo actualizados.
  async function guardarFotos() {
    if (!fotosCambiadas || !fotos.length) return {};
    if (fotos.some(f => !f.blob)) throw new Error('Esperá a que terminen de cargar las fotos.');
    for (const f of fotos.filter(x => x.nueva)) {
      f.archivoId = await estado.api.subirFotoOriginal({ ...r, clienteNombre: cli.nombre }, f.blob);
      f.nueva = false;
    }
    const pdf = await armarPdf(fotos);
    const arch = await estado.api.reemplazarPdf({ ...r, clienteNombre: cli.nombre }, pdf);
    Object.assign(r, arch);
    r.paginas = fotos.map(f => ({ archivoId: f.archivoId, giro: f.giro || 0, recorte: f.recorte || null, mejorar: !!f.mejorar }));
    fotos.forEach(f => { f.cambiada = false; });
    fotosCambiadas = false;
    cargarVisor(pdf);
    pintarFotos();
    return { ...arch, paginas: r.paginas };
  }

  $('#btn-guardar-foto', vista).onclick = async () => {
    ocupado(true);
    try {
      const cambios = await guardarFotos();
      await estado.api.actualizar({ ...leer(), ...cambios });
      aviso('Foto actualizada.', 'ok');
    } catch (e) { msg(esc(e.message), 'error'); }
    finally { ocupado(false); ayudaFotos(); }
  };

  function leer() {
    return {
      ...r,
      remito: normalizarRemito($('#f-pv', vista).value, $('#f-numero', vista).value),
      fechaRemito: $('#f-fecha', vista).value,
      tecnico: $('#f-tecnico', vista).value.trim(),
      fechaSolicitud: $('#f-solicitud', vista).value,
      periodo: $('#f-periodo', vista).value,
      establecimiento: $('#f-estab', vista).value,
      items: grupos.flatMap((gr, g) => cli.items.filter(it => +gr.cant[it.cod] > 0)
        .map(it => ({ cod: it.cod, servicio: it.servicio, cant: +gr.cant[it.cod], g }))),
      km: grupos.reduce((s, gr) => s + (+gr.cant[cli.items.find(it => it.unidad === 'km')?.cod ?? 'KM'] || 0), 0),
      extras: leerExtras(),
      gruposExtras: conGrupos ? grupos.map(gr => valoresDeGrupo(cli.extras, gr.extras)) : [],
      revisor: estado.usuario.displayName,
    };
  }

  const msg = (t, tipo) => { $('#msg', vista).innerHTML = t ? `<div class="mensaje ${tipo}">${t}</div>` : ''; };
  const ocupado = on => vista.querySelectorAll('button').forEach(b => { b.disabled = on; });

  $('#btn-guardar', vista).onclick = async () => {
    ocupado(true);
    try { const f = await guardarFotos(); await estado.api.actualizar({ ...leer(), ...f }); aviso('Cambios guardados.', 'ok'); }
    catch (e) { msg(esc(e.message), 'error'); }
    finally { ocupado(false); }
  };

  $('#btn-eliminar', vista).onclick = async () => {
    if (!confirm(`¿Eliminar el remito ${r.remito} (${r.establecimiento}), su PDF y sus fotos?\n\n` +
      'Va a la papelera de reciclaje del sitio de SharePoint: se puede recuperar durante 93 días.')) return;
    ocupado(true);
    try {
      await estado.api.eliminarRemito(r);
      aviso(`Remito ${r.remito} eliminado.`, 'ok');
      ir('revision');
    } catch (e) { msg(esc(e.message), 'error'); ocupado(false); }
  };

  $('#btn-rechazar', vista).onclick = async () => {
    const comentario = $('#f-comentario', vista).value.trim();
    if (!comentario) { msg('Escribí el motivo del rechazo para el técnico.', 'error'); return; }
    ocupado(true);
    try {
      await estado.api.actualizar({ ...leer(), estado: 'Rechazado', comentario, fechaRevision: new Date().toISOString() });
      aviso('Remito rechazado. El técnico lo ve en "Mis remitos".', 'ok');
      ir('revision');
    } catch (e) { msg(esc(e.message), 'error'); ocupado(false); }
  };

  $('#btn-aprobar', vista).onclick = async () => {
    const d = leer();
    const faltan = [];
    if (!d.remito) faltan.push('número de remito');
    if (!d.fechaRemito) faltan.push('fecha del remito');
    if (!d.fechaSolicitud) faltan.push('fecha de solicitud');
    if (!d.periodo) faltan.push('período');
    if (!d.tecnico) faltan.push('responsable');
    for (const x of extrasRemito.filter(x => x.quien !== 'auto' && x.obligatorio)) {
      if (!d.extras[x.campo]) faltan.push(x.etiqueta.toLowerCase());
    }
    for (const x of camposDeGrupo(cli.extras).filter(x => x.quien !== 'auto' && x.obligatorio)) {
      d.gruposExtras.forEach((ge, g) => { if (!ge[x.campo]) faltan.push(`${x.etiqueta.toLowerCase()}${grupos.length > 1 ? ` (grupo ${g + 1})` : ''}`); });
    }
    if (!d.items.some(i => cli.items.find(c => String(c.cod) === String(i.cod))?.planilla !== false)) faltan.push('al menos un servicio');
    if (faltan.length) { msg('Falta: ' + faltan.join(', ') + '.', 'error'); return; }
    if (!cli.excel?.itemId && estado.api.modo !== 'demo') { msg('Este cliente no tiene Excel vinculado (Configuración).', 'error'); return; }

    ocupado(true);
    msg('Escribiendo en el Excel…', 'alerta');
    recordar('ultimaFechaSolicitud', d.fechaSolicitud);
    recordar('ultimoPeriodo', d.periodo);
    const revision = { ...d, fechaRevision: new Date().toISOString() };
    try {
      // Primero se guardan las correcciones; si el Excel falla, queda "Pendiente Excel" para reintentar
      // Si se retocó la foto, el PDF nuevo es el que se archiva junto al Excel
      Object.assign(revision, await guardarFotos());
      await estado.api.actualizar({ ...revision, estado: 'Pendiente Excel', comentario: 'Pasando al Excel…' });
      const filas = await estado.api.escribirExcel(cli, revision);
      // El PDF va a la carpeta del cliente; si falla, el remito igual queda aprobado (el Excel ya se escribió)
      let archivoUrl = revision.archivoUrl, comentario = '';
      try {
        archivoUrl = await estado.api.archivarPdf(cli, revision);
      } catch (err) {
        comentario = 'Aprobado, pero el PDF no se pudo copiar a la carpeta del cliente: ' + err.message;
      }
      await estado.api.actualizar({ ...revision, estado: 'Aprobado', comentario, filasExcel: filas, archivoUrl });
      if (comentario) aviso(comentario, 'error');
      aviso(`Aprobado. Filas ${filas} de la planilla${comentario ? '' : ' y PDF guardado junto al Excel'}.`, 'ok');
      ir('revision');
    } catch (e) {
      console.error(e);
      await estado.api.actualizar({ ...revision, estado: 'Pendiente Excel', comentario: e.message }).catch(() => {});
      msg(`<b>No se pasó al Excel:</b> ${esc(e.message)}` +
        (e.codigo === 'duplicado' ? '<div class="acciones"><button id="btn-ya-esta">Ya está en la planilla: marcar aprobado</button></div>' : ''), 'error');
      ocupado(false);
      $('#btn-ya-esta', vista)?.addEventListener('click', async () => {
        const archivoUrl = await estado.api.archivarPdf(cli, revision).catch(() => revision.archivoUrl);
        await estado.api.actualizar({ ...revision, estado: 'Aprobado', comentario: 'Ya estaba en la planilla (no se escribió).', archivoUrl });
        aviso('Marcado como aprobado.', 'ok');
        ir('revision');
      });
    }
  };
}
