(function () {
  const MESES_ES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  function formatearMesAnio(fechaIso) {
    if (!fechaIso) return '';
    const d = new Date(fechaIso);
    if (isNaN(d.getTime())) return fechaIso;
    return `${MESES_ES[d.getMonth()]} de ${d.getFullYear()}`;
  }

  // Hoja de estilos del modal, inyectada UNA sola vez. Usa las mismas
  // variables de color (--accent, --bg-card, --border, etc.) que ya
  // define cada página -- así el modal siempre queda fiel al estilo de
  // la página donde se abre, sin tener que repetir colores a mano.
  function inyectarEstilos() {
    if (document.getElementById('audit-ui-modal-styles')) return;
    const style = document.createElement('style');
    style.id = 'audit-ui-modal-styles';
    style.textContent = `
      .aui-overlay {
        position: fixed; inset: 0; background: rgba(15,23,42,.48);
        display: flex; align-items: center; justify-content: center;
        z-index: 9999; padding: 16px;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
      }
      .aui-box {
        background: var(--bg-card, #fff); color: var(--text-primary, #1e293b);
        width: min(620px, 100%); max-height: 90vh; overflow: auto;
        border-radius: 14px; padding: 22px;
        box-shadow: 0 24px 70px rgba(15,23,42,.28);
        font-family: inherit;
      }
      .aui-box * { font-family: inherit; box-sizing: border-box; }
      .aui-title { margin: 0 0 6px; font-size: 1.05rem; font-weight: 700; color: var(--text-primary, #1e293b); }
      .aui-subtitle { margin: 0 0 16px; color: var(--text-secondary, #64748b); font-size: 0.85rem; line-height: 1.5; }
      .aui-radio-row {
        display: flex; align-items: center; gap: 10px; padding: 10px 12px;
        font-weight: 600; font-size: 0.85rem; border-radius: 9px;
        border: 1px solid var(--border, #e2e8f0); margin-bottom: 8px; cursor: pointer;
        transition: border-color .15s, background .15s;
      }
      .aui-radio-row:hover { border-color: var(--accent, #4f46e5); }
      .aui-radio-row input[type=radio] { accent-color: var(--accent, #4f46e5); width: 16px; height: 16px; margin: 0; cursor: pointer; }
      .aui-group-label { font-weight: 700; font-size: 0.8rem; color: var(--accent, #4f46e5); margin: 4px 0 6px; text-transform: uppercase; letter-spacing: .03em; }
      .aui-list {
        border: 1px solid var(--border, #e2e8f0); border-radius: 10px;
        padding: 10px; margin: 6px 0 16px; max-height: 220px; overflow: auto;
      }
      .aui-list-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
      .aui-list-header strong { font-size: 0.85rem; color: var(--text-primary, #1e293b); }
      .aui-list-count { color: var(--text-secondary, #64748b); font-size: 0.75rem; }
      .aui-checkbox-row {
        display: flex; align-items: center; gap: 8px; padding: 6px 4px;
        font-size: 0.82rem; border-radius: 6px; cursor: pointer;
      }
      .aui-checkbox-row:hover { background: var(--bg-body, #f5f7fc); }
      .aui-checkbox-row input[type=checkbox] { accent-color: var(--accent, #4f46e5); width: 15px; height: 15px; margin: 0; cursor: pointer; }
      .aui-checkbox-row small { margin-left: auto; color: var(--text-secondary, #64748b); font-size: 0.7rem; white-space: nowrap; }
      .aui-radio-item {
        display: flex; align-items: center; gap: 9px; padding: 8px 6px;
        cursor: pointer; border-radius: 7px; font-size: 0.85rem;
      }
      .aui-radio-item:hover { background: var(--bg-body, #f5f7fc); }
      .aui-radio-item input[type=radio] { accent-color: var(--accent, #4f46e5); width: 16px; height: 16px; margin: 0; cursor: pointer; }
      .aui-btn {
        font-size: 0.78rem; font-weight: 700; border-radius: 8px; padding: 8px 14px;
        cursor: pointer; transition: all .15s; border: 1px solid var(--border, #e2e8f0);
        background: var(--bg-card, #fff); color: var(--text-primary, #1e293b);
      }
      .aui-btn:hover { border-color: var(--accent, #4f46e5); color: var(--accent, #4f46e5); }
      .aui-btn-ghost-sm {
        font-size: 0.7rem; font-weight: 700; border-radius: 7px; padding: 6px 10px;
        cursor: pointer; border: 1px solid var(--border, #e2e8f0);
        background: var(--bg-body, #f5f7fc); color: var(--text-secondary, #64748b);
      }
      .aui-btn-ghost-sm:hover { border-color: var(--accent, #4f46e5); color: var(--accent, #4f46e5); }
      .aui-btn-primary {
        font-size: 0.8rem; font-weight: 700; border-radius: 9px; padding: 10px 20px;
        cursor: pointer; border: none; color: #fff;
        background: linear-gradient(135deg, var(--success, #10b981), #059669);
        transition: all .15s;
      }
      .aui-btn-primary:hover { opacity: .92; transform: translateY(-1px); }
      .aui-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 18px; }
    `;
    document.head.appendChild(style);
  }

  window.solicitarDecisionDuplicados = function (data) {
    inyectarEstilos();
    const previas = Object.keys(data.ya_descargadas || {});
    if (!previas.length) return Promise.resolve(null);
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'aui-overlay';
      const box = document.createElement('div');
      box.className = 'aui-box';
      box.innerHTML = `<h3 class="aui-title">Facturas encontradas previamente</h3><p class="aui-subtitle">Se encontraron <strong>${previas.length}</strong> facturas ya descargadas <strong>este mismo mes</strong>. Elige qué hacer antes de continuar.</p>`;
      const options = [
        ['ninguna', 'No volver a descargar ninguna'],
        ['todas', 'Volver a descargar todas'],
        ['seleccionadas', 'Seleccionar cuáles volver a descargar']
      ];
      const radios = options.map(([value, label], index) => `<label class="aui-radio-row"><input type="radio" name="duplicateDecision" value="${value}" ${index === 0 ? 'checked' : ''}>${label}</label>`).join('');
      box.insertAdjacentHTML('beforeend', radios);
      const list = document.createElement('div');
      list.className = 'aui-list';
      list.style.display = 'none';
      list.innerHTML = `<div class="aui-list-header"><strong>Facturas detectadas</strong><span id="duplicateCount" class="aui-list-count"></span></div><div style="display:flex;gap:8px;margin-bottom:8px"><button type="button" class="aui-btn-ghost-sm" data-select="all">Seleccionar todas</button><button type="button" class="aui-btn-ghost-sm" data-select="none">Deseleccionar todas</button></div>` + previas.map(f => `<label class="aui-checkbox-row"><input type="checkbox" value="${f}"> <span>${f}</span><small>Descargada en ${formatearMesAnio(data.ya_descargadas[f])}</small></label>`).join('');
      box.appendChild(list);
      box.insertAdjacentHTML('beforeend', '<div class="aui-footer"><button type="button" class="aui-btn" data-cancel>Cancelar</button><button type="button" class="aui-btn-primary" data-continue>Continuar</button></div>');
      const update = () => { const chosen = list.querySelectorAll('input[type=checkbox]:checked').length; const count = list.querySelector('#duplicateCount'); if (count) count.textContent = `${chosen} seleccionadas`; };
      box.querySelectorAll('input[name=duplicateDecision]').forEach(r => r.addEventListener('change', e => { list.style.display = e.target.value === 'seleccionadas' ? 'block' : 'none'; update(); }));
      list.querySelector('[data-select=all]').onclick = () => { list.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = true); update(); };
      list.querySelector('[data-select=none]').onclick = () => { list.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = false); update(); };
      list.addEventListener('change', update);
      box.querySelector('[data-cancel]').onclick = () => { overlay.remove(); resolve(null); };
      box.querySelector('[data-continue]').onclick = () => { const decision = box.querySelector('input[name=duplicateDecision]:checked').value; const selected = [...list.querySelectorAll('input[type=checkbox]:checked')].map(c => c.value); overlay.remove(); resolve({ decision, selected }); };
      overlay.appendChild(box); document.body.appendChild(overlay);
    });
  };

  window.solicitarSeleccionIPS = function (catalogoIps, identidadActual) {
    inyectarEstilos();
    return new Promise(resolve => {
      const mapaIdentidad = { "SaludNet": "Salud Net", "Campbell": "Campbell" };
      const soloGrupo = mapaIdentidad[identidadActual];
      const overlay = document.createElement('div');
      overlay.className = 'aui-overlay';
      const box = document.createElement('div');
      box.className = 'aui-box';
      box.style.width = 'min(560px, 100%)';
      box.innerHTML = `<h3 class="aui-title">No se identificó la IPS automáticamente</h3><p class="aui-subtitle">Esta cuenta es nueva. Selecciona a cuál IPS corresponde para continuar.</p>`;
      const grupoHtml = (titulo, lista) => {
        if (!lista.length) return '';
        const items = lista.map(ips => `<label class="aui-radio-item"><input type="radio" name="ipsSeleccionada" value="${ips.nit}"><span>${ips.nombre_estandar}</span></label>`).join('');
        return `<div style="margin-bottom:14px"><div class="aui-group-label">${titulo}</div>${items}</div>`;
      };
      const grupos = soloGrupo
        ? grupoHtml(soloGrupo, catalogoIps[soloGrupo] || [])
        : grupoHtml('Salud Net', catalogoIps['Salud Net'] || []) + grupoHtml('Campbell', catalogoIps['Campbell'] || []);
      box.insertAdjacentHTML('beforeend', grupos);
      box.insertAdjacentHTML('beforeend', '<div class="aui-footer"><button type="button" class="aui-btn" data-cancel>Cancelar</button><button type="button" class="aui-btn-primary" data-continue>Continuar</button></div>');
      box.querySelector('[data-cancel]').onclick = () => { overlay.remove(); resolve(null); };
      box.querySelector('[data-continue]').onclick = () => {
        const elegido = box.querySelector('input[name=ipsSeleccionada]:checked');
        if (!elegido) return;
        overlay.remove(); resolve(elegido.value);
      };
      overlay.appendChild(box); document.body.appendChild(overlay);
    });
  };
})();
