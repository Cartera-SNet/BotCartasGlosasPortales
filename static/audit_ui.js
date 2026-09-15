(function () {
  window.solicitarDecisionDuplicados = function (data) {
    const previas = Object.keys(data.ya_descargadas || {});
    if (!previas.length) return Promise.resolve(null);
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.48);display:flex;align-items:center;justify-content:center;z-index:9999;padding:16px';
      const box = document.createElement('div');
      box.style.cssText = 'background:#fff;color:#172033;width:min(620px,100%);max-height:90vh;overflow:auto;border-radius:14px;padding:22px;box-shadow:0 24px 70px rgba(15,23,42,.28);font-family:inherit';
      box.innerHTML = `<h3 style="margin:0 0 6px">Facturas encontradas previamente</h3><p style="margin:0 0 16px;color:#64748b">Se encontraron <strong>${previas.length}</strong> facturas procesadas anteriormente. Elige qué hacer antes de continuar.</p>`;
      const options = [
        ['ninguna', 'No volver a descargar ninguna'],
        ['todas', 'Volver a descargar todas'],
        ['seleccionadas', 'Seleccionar cuáles volver a descargar']
      ];
      const radios = options.map(([value, label], index) => `<label style="display:flex;gap:9px;align-items:center;padding:9px 0;font-weight:600"><input type="radio" name="duplicateDecision" value="${value}" ${index === 0 ? 'checked' : ''}>${label}</label>`).join('');
      box.insertAdjacentHTML('beforeend', radios);
      const list = document.createElement('div');
      list.style.cssText = 'display:none;border:1px solid #e2e8f0;border-radius:8px;padding:9px;margin:6px 0 14px;max-height:220px;overflow:auto';
      list.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:7px"><strong>Facturas detectadas</strong><span id="duplicateCount" style="color:#64748b;font-size:.85rem"></span></div><div style="display:flex;gap:8px;margin-bottom:8px"><button type="button" data-select="all">Seleccionar todas</button><button type="button" data-select="none">Deseleccionar todas</button></div>` + previas.map(f => `<label style="display:flex;gap:8px;padding:5px 0;font-size:.9rem"><input type="checkbox" value="${f}"> <span>${f}</span><small style="margin-left:auto;color:#64748b">${data.ya_descargadas[f] || ''}</small></label>`).join('');
      box.appendChild(list);
      box.insertAdjacentHTML('beforeend', '<div style="display:flex;justify-content:flex-end;gap:8px;margin-top:16px"><button type="button" data-cancel>Cancelar</button><button type="button" data-continue style="background:#2563eb;color:#fff;border:0;border-radius:7px;padding:9px 15px;font-weight:700">Continuar</button></div>');
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
})();
