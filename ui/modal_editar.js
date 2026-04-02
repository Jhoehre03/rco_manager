// ── MODAL: VER / EDITAR NOTAS JÁ LANÇADAS ────────────────────────────────────
// Usado quando a AV já foi lançada.
// Colunas: Nº | Nome | Planilha | RCO atual (travado) | Enviar (editável)
// Sem seleção de data/valor. Busca notas do RCO via API.

let _me_turma  = null;
let _me_alunos = [];

// ── HTML do modal ─────────────────────────────────────────────────────────────

function _meHTML() {
  return `
    <div id="modal-editar" class="modal-overlay hidden">
      <div class="modal-box" style="max-width:680px;width:95vw">
        <div class="modal-header">
          <span id="me-titulo">Ver / Editar Notas</span>
          <button class="modal-close" onclick="_meFechar()">✕</button>
        </div>
        <div class="modal-body" style="padding:18px">
          <div style="margin-bottom:12px;display:flex;align-items:center;gap:10px">
            <div id="me-av-info" style="font-size:13px;font-weight:600;color:#1e293b">—</div>
            <span id="me-data-lanc" style="background:#dcfce7;color:#166534;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600"></span>
          </div>
          <div id="me-aviso" style="display:none;font-size:12px;margin-bottom:10px"></div>
          <div id="me-tabela" class="hidden">
            <table style="width:100%;border-collapse:collapse;font-size:13px">
              <thead>
                <tr style="background:#f8fafc">
                  <th style="padding:6px 8px;text-align:left;width:36px">Nº</th>
                  <th style="padding:6px 8px;text-align:left">Nome</th>
                  <th style="padding:6px 8px;text-align:center;width:80px">Planilha</th>
                  <th style="padding:6px 8px;text-align:center;width:80px">RCO atual</th>
                  <th style="padding:6px 8px;text-align:center;width:90px">Enviar</th>
                </tr>
              </thead>
              <tbody id="me-tbody"></tbody>
            </table>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn" onclick="_meFechar()">Fechar</button>
          <button class="btn btn-sm" id="me-btn-buscar-planilha" onclick="_meBuscarPlanilha()" style="background:#e0e7ff;color:#3730a3">📋 Buscar da planilha</button>
          <button class="btn btn-sm" id="me-btn-buscar-rco" onclick="_meBuscarRco()" style="background:#f1f5f9;color:#374151" disabled>🔍 Ver notas RCO</button>
          <button class="btn btn-primary" id="me-btn-editar" disabled onclick="_meEditar()">Editar no RCO</button>
        </div>
      </div>
    </div>`;
}

// ── Abrir ─────────────────────────────────────────────────────────────────────

function _meAbrir(turmaIdx, tri, nomeAv, dataLancamento) {
  _me_turma  = state.turmas[turmaIdx];
  _me_alunos = [];
  window._me_tri = tri;
  window._me_av  = nomeAv;

  document.getElementById('me-titulo').textContent    = `${nomeAv} — ${_me_turma.turma}`;
  document.getElementById('me-av-info').textContent   = `${nomeAv} · ${tri}º Tri`;
  document.getElementById('me-data-lanc').textContent = dataLancamento ? `✓ Lançada em ${dataLancamento}` : '';
  document.getElementById('me-aviso').style.display   = 'none';
  document.getElementById('me-tabela').classList.add('hidden');
  document.getElementById('me-tbody').innerHTML       = '';
  document.getElementById('me-btn-buscar-rco').disabled = true;
  document.getElementById('me-btn-editar').disabled   = true;
  document.getElementById('modal-editar').classList.remove('hidden');
}

function _meFechar() {
  document.getElementById('modal-editar').classList.add('hidden');
}

// ── Buscar notas da planilha ──────────────────────────────────────────────────

async function _meBuscarPlanilha() {
  const tri    = window._me_tri || 1;
  const coluna = window._me_av  || '';
  if (!coluna) { toast('AV não definida', 'error'); return; }

  const btn = document.getElementById('me-btn-buscar-planilha');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Buscando...';

  try {
    const {escola, turma, disciplina} = _me_turma;
    const res = await window.pywebview.api.get_notas_planilha(escola, turma, disciplina, tri, coluna);
    if (!res.ok) { toast('Erro: ' + res.erro, 'error'); return; }

    _me_alunos = res.alunos.map(a => ({...a, nota_rco_atual: null}));
    _meRenderTabela();
    document.getElementById('me-tabela').classList.remove('hidden');
    document.getElementById('me-btn-buscar-rco').disabled = !state.status?.api_conectada;
    document.getElementById('me-btn-editar').disabled = !state.chrome;
  } catch(e) {
    toast('Erro inesperado: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '📋 Buscar da planilha';
  }
}

// ── Buscar notas do RCO via API ───────────────────────────────────────────────

async function _meBuscarRco() {
  if (!state.status?.api_conectada) { toast('API não conectada', 'error'); return; }
  if (_me_alunos.length === 0) { toast('Busque as notas da planilha primeiro', 'error'); return; }

  const tri    = window._me_tri || 1;
  const tipoAv = window._me_av  || '';

  const btn = document.getElementById('me-btn-buscar-rco');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="border-top-color:#3730a3;border-color:rgba(55,48,163,.2)"></span>Buscando...';

  try {
    const {escola, turma, disciplina} = _me_turma;
    const res = await window.pywebview.api.get_notas_rco_api(
      escola, turma, disciplina, `${tri}T`, tipoAv
    );
    if (!res.ok) { toast('Erro: ' + res.erro, 'error'); return; }

    const rcoIdx = {};
    res.alunos.forEach(a => { rcoIdx[a.numero] = a.nota; });
    _me_alunos.forEach(a => { a.nota_rco_atual = rcoIdx[a.numero] ?? null; });

    _meRenderTabela();
    toast('Notas do RCO carregadas.', 'success');
  } catch(e) {
    toast('Erro inesperado: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '🔍 Ver notas RCO';
  }
}

// ── Renderizar tabela ─────────────────────────────────────────────────────────

function _meRenderTabela() {
  document.getElementById('me-tbody').innerHTML = _me_alunos.map((a, i) => {
    const inBadge = a.inativo
      ? '<span style="font-size:10px;background:#fee2e2;color:#991b1b;padding:1px 6px;border-radius:10px;margin-left:6px">Inativo</span>'
      : '';

    const notaExibida = a.nota_rco && a.nota_rco !== '0'
      ? (parseFloat(a.nota_rco) / 10).toFixed(1).replace('.', ',')
      : '—';

    const rcoAtual = a.nota_rco_atual != null
      ? `<span style="color:#6366f1;font-weight:500">${a.nota_rco_atual}</span>`
      : '<span style="color:#cbd5e1">—</span>';

    const diverge = a.nota_rco_atual != null && a.nota_rco_atual !== '-'
      && a.nota_rco && a.nota_rco !== a.nota_rco_atual;
    const rowBg = diverge ? 'background:#fffbeb' : '';

    return `<tr style="${a.inativo ? 'opacity:.5;' : ''}${rowBg}">
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;color:#94a3b8">${a.numero}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9">${a.nome}${inBadge}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center;color:#64748b">${notaExibida}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center">${rcoAtual}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center">
        <input type="text" maxlength="3" ${a.inativo ? 'disabled' : ''}
               value="${a.nota_rco && a.nota_rco !== '0' ? a.nota_rco : ''}"
               oninput="_me_alunos[${i}].nota_rco=this.value"
               style="width:60px;text-align:center;border:1px solid #e2e8f0;border-radius:6px;
                      padding:3px 6px;font-size:13px;background:${a.inativo?'#f8fafc':'#fff'}">
      </td>
    </tr>`;
  }).join('');
}

// ── Editar notas ──────────────────────────────────────────────────────────────

async function _meEditar() {
  if (!state.chrome) { toast('Conecte o Chrome primeiro', 'error'); return; }

  const tri    = window._me_tri || 0;
  const tipoAv = window._me_av  || '';
  if (!tri || !tipoAv) { toast('AV não definida', 'error'); return; }

  const notas = _me_alunos
    .filter(a => !a.inativo && a.nota_rco && a.nota_rco !== '0')
    .map(a => ({
      nome_normalizado: a.nome.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim(),
      nota: a.nota_rco,
    }));

  if (notas.length === 0) { toast('Nenhuma nota para enviar', 'error'); return; }

  const triStr = tri === 1 ? '1º Tri' : tri === 2 ? '2º Tri' : '3º Tri';
  const btn    = document.getElementById('me-btn-editar');
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span>Editando ${notas.length} notas...`;

  try {
    const {escola, turma, disciplina} = _me_turma;
    const res = await window.pywebview.api.editar_notas(
      escola, turma, disciplina, triStr, tipoAv, notas
    );
    if (res.ok) {
      toast('Notas editadas com sucesso!', 'success');
      _meFechar();
    } else {
      toast('Erro: ' + res.erro, 'error');
    }
  } catch(e) {
    toast('Erro inesperado: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Editar no RCO';
  }
}
