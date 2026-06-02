import re

content = open('templates/mantenimiento_hr.html', encoding='utf-8').read()

js_asignar = r"""
// ── Asignar SAP ───────────────────────────────────────────────────────────────
let _planAsignacion = null;
let _planHrOrigen   = null;

async function abrirPlanAsignacion(id_hruta) {
  _planHrOrigen = id_hruta;
  const modal = document.getElementById("modal-asignar");
  const body  = document.getElementById("modal-asignar-body");
  modal.style.display = "flex";
  body.innerHTML = '<div style="text-align:center;padding:2rem;color:var(--muted);"><span style="animation:spin 1s linear infinite;display:inline-block;font-size:1.5rem;">&#x27F3;</span><br><br>Calculando plan...</div>';
  try {
    const r = await fetch("/api/mantenimiento_hr/plan_asignacion/" + id_hruta);
    const d = await r.json();
    if (!d.ok) { body.innerHTML = '<div style="color:var(--err-txt);padding:1rem;">' + (d.error||"Error") + '</div>'; return; }
    _planAsignacion = d;
    _renderPlan(d);
  } catch(e) {
    body.innerHTML = '<div style="color:var(--err-txt);padding:1rem;">Error: ' + e.message + '</div>';
  }
}

function _renderPlan(d) {
  const body = document.getElementById("modal-asignar-body");
  const sinHrHtml = (d.sin_hr && d.sin_hr.length)
    ? '<div style="margin-top:.8rem;padding:.5rem .75rem;background:rgba(234,179,8,.1);border:1px solid rgba(234,179,8,.3);border-radius:8px;font-size:.78rem;color:#ca8a04;"><i class="bi bi-exclamation-triangle me-1"></i><b>' + d.sin_hr.length + ' ZFERs sin HR disponible</b>: ' + d.sin_hr.slice(0,5).join(", ") + (d.sin_hr.length>5?" ...":"") + '</div>'
    : '';
  const batchesHtml = d.batches.map((b,i) =>
    '<div style="background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:.7rem .9rem;margin-bottom:.5rem;">' +
    '<div style="display:flex;align-items:center;gap:.6rem;flex-wrap:wrap;">' +
    '<span style="font-family:monospace;font-weight:700;color:var(--accent);">' + b.hr_destino + '</span>' +
    '<span style="font-size:.82rem;color:var(--text);flex:1;">' + (b.hr_desc||"") + '</span>' +
    '<span style="font-size:.75rem;color:var(--muted);">' + b.materiales_actuales + ' actuales</span>' +
    '<span style="background:rgba(34,197,94,.12);color:var(--ok-txt);border:1px solid var(--ok-bdr);border-radius:6px;padding:.15rem .55rem;font-size:.75rem;font-weight:700;">+' + b.n_zfers + ' ZFERs</span></div>' +
    '<div id="res-batch-' + i + '" style="margin-top:.3rem;font-size:.78rem;display:none;"></div></div>'
  ).join("");
  body.innerHTML =
    '<div style="font-size:.82rem;color:var(--muted);margin-bottom:.8rem;">' + d.total_zfers + ' ZFERs fuera &rarr; ' + d.asignables + ' asignables en ' + d.batches.length + ' HR(s)</div>' +
    batchesHtml + sinHrHtml;
}

async function ejecutarPlanAsignacion() {
  if (!_planAsignacion || !_planAsignacion.batches.length) return;
  const btnEj = document.getElementById("btn-ejecutar-asignacion");
  btnEj.disabled = true;
  btnEj.innerHTML = '<span style="animation:spin 1s linear infinite;display:inline-block;">&#x27F3;</span> Ejecutando...';
  let okCount = 0, errCount = 0;

  for (let i = 0; i < _planAsignacion.batches.length; i++) {
    const b = _planAsignacion.batches[i];
    const resDiv = document.getElementById("res-batch-" + i);
    if (resDiv) { resDiv.style.display="block"; resDiv.innerHTML='<span style="color:var(--muted);">&#x27F3; Asignando...</span>'; }
    try {
      const r = await fetch("/api/mantenimiento_hr/ejecutar_asignacion", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({batch: b})
      });
      const d = await r.json();
      if (d.ok) {
        okCount++;
        if (resDiv) resDiv.innerHTML = '<span style="color:var(--ok-txt);"><i class="bi bi-check-circle me-1"></i>' + (d.mensaje||"OK") + '</span>';
      } else {
        errCount++;
        const log = (d.detalles && d.detalles.length) ? " | " + d.detalles.slice(-2).join(" | ") : "";
        if (resDiv) resDiv.innerHTML = '<span style="color:var(--err-txt);"><i class="bi bi-x-circle me-1"></i>' + (d.error||"Error") + log + '</span>';
      }
    } catch(e) {
      errCount++;
      if (resDiv) resDiv.innerHTML = '<span style="color:var(--err-txt);">Error: ' + e.message + '</span>';
    }
  }

  btnEj.disabled = false;
  btnEj.innerHTML = okCount + ' OK' + (errCount ? ' | ' + errCount + ' errores' : '') + ' — cerrar';
  btnEj.onclick = () => document.getElementById("modal-asignar").style.display="none";

  // Panel inline en la card
  const panelId = "resultado-asig-" + _planHrOrigen;
  let panel = document.getElementById(panelId);
  if (!panel) {
    const card = document.getElementById("btn-asignar-" + _planHrOrigen);
    if (card) {
      panel = document.createElement("div");
      panel.id = panelId;
      card.closest(".card").appendChild(panel);
    }
  }
  if (panel) {
    const ok = errCount === 0;
    panel.style.cssText = "margin-top:.5rem;padding:.6rem 1rem;border-radius:8px;font-size:.82rem;display:block;background:" +
      (ok?"rgba(34,197,94,.1)":"rgba(239,68,68,.1)") + ";border:1px solid " +
      (ok?"var(--ok-bdr)":"var(--err-bdr)") + ";color:" + (ok?"var(--ok-txt)":"var(--err-txt)");
    panel.innerHTML = (ok?'<i class="bi bi-check-circle me-1"></i>':'<i class="bi bi-exclamation-circle me-1"></i>') +
      "<b>" + okCount + " batches asignados" + (errCount?" | "+errCount+" errores":"") + "</b>";
  }
}
"""

modal_asignar = """
<!-- Modal plan asignacion -->
<div id="modal-asignar" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,.75);
     z-index:9999; align-items:center; justify-content:center;">
  <div style="background:var(--card-bg); border:1px solid var(--border); border-radius:16px;
              padding:1.8rem; max-width:560px; width:95%; max-height:80vh; display:flex; flex-direction:column;">
    <h5 style="margin:0 0 .4rem; font-weight:800; font-size:1rem;">
      <i class="bi bi-plus-circle me-2" style="color:#16a34a;"></i>Plan de Asignacion SAP
    </h5>
    <div style="color:var(--muted); font-size:.82rem; margin-bottom:1rem;">
      Se busca la HR adecuada para cada ZFER respetando el limite de 300 materiales.
    </div>
    <div id="modal-asignar-body" style="flex:1; overflow-y:auto; margin-bottom:1rem;"></div>
    <div style="display:flex; gap:.6rem;">
      <button id="btn-ejecutar-asignacion" onclick="ejecutarPlanAsignacion()"
        style="flex:1; background:#16a34a; color:#fff; border:none; border-radius:10px;
               padding:.6rem; font-size:.88rem; font-weight:700; cursor:pointer;">
        <i class="bi bi-play-circle me-1"></i>Ejecutar asignacion
      </button>
      <button onclick="document.getElementById('modal-asignar').style.display='none'"
        style="background:none; border:1px solid var(--border); color:var(--muted);
               border-radius:10px; padding:.6rem 1rem; font-size:.85rem; cursor:pointer;">
        Cancelar
      </button>
    </div>
  </div>
</div>

"""

# Insert JS before last </script>
last_endblock = content.rfind('{% endblock %}')
last_script_close = content.rfind('</script>', 0, last_endblock)
content = content[:last_script_close] + js_asignar + content[last_script_close:]

# Insert modal before last {% endblock %}
last_endblock = content.rfind('{% endblock %}')
content = content[:last_endblock] + modal_asignar + content[last_endblock:]

open('templates/mantenimiento_hr.html', 'w', encoding='utf-8').write(content)
print('OK')
