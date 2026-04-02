// ── MODAL: NOVO LANÇAMENTO DE NOTAS ──────────────────────────────────────────
// Usado quando a AV ainda não foi lançada.
// Colunas: Nº | Nome | Planilha | Enviar (editável)
// Sem coluna "RCO atual", sem modo editar.

let _ml_turma  = null;
let _ml_alunos = [];

// ── HTML do modal ─────────────────────────────────────────────────────────────

function _mlHTML() {
  return `
    <div id="modal-lancar" class="modal-overlay hidden">
      <div class="modal-box" style="max-width:680px;width:95vw">
        <div class="modal-header">
          <span id="ml-titulo">Lançar Notas</span>
          <button class="modal-close" onclick="_mlFechar()">✕</button>
        </div>
        <div class="modal-body" style="padding:18px">
          <div style="margin-bottom:8px">
            <div id="ml-av-info" style="font-size:13px;font-weight:600;color:#1e293b">—</div>
          </div>
          <div class="form-row-2" style="margin-bottom:12px">
            <div>
              <label class="form-label">Data</label>
              <select id="ml-data" class="form-input">
                <option value="">Carregando...</option>
              </select>
            </div>
            <div>
              <label class="form-label">Valor máximo</label>
              <input type="number" id="ml-valor" class="form-input" step="0.1" min="0" max="100" placeholder="ex: 3.0">
            </div>
          </div>
          <div id="ml-aviso" style="display:none;font-size:12px;margin-bottom:10px"></div>
          <div id="ml-tabela" class="hidden">
            <table style="width:100%;border-collapse:collapse;font-size:13px">
              <thead>
                <tr style="background:#f8fafc">
                  <th style="padding:6px 8px;text-align:left;width:36px">Nº</th>
                  <th style="padding:6px 8px;text-align:left">Nome</th>
                  <th style="padding:6px 8px;text-align:center;width:80px">Planilha</th>
                  <th style="padding:6px 8px;text-align:center;width:90px">Enviar</th>
                </tr>
              </thead>
              <tbody id="ml-tbody"></tbody>
            </table>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn" onclick="_mlFechar()">Fechar</button>
          <button class="btn btn-sm" id="ml-btn-buscar" onclick="_mlBuscarPlanilha()" style="background:#e0e7ff;color:#3730a3">📋 Buscar da planilha</button>
          <button class="btn btn-primary" id="ml-btn-lancar" disabled onclick="_mlLancar()">Lançar no RCO</button>
        </div>
      </div>
    </div>`;
}

// ── Abrir ─────────────────────────────────────────────────────────────────────

function _mlAbrir(turmaIdx, tri, nomeAv) {
  _ml_turma  = state.turmas[turmaIdx];
  _ml_alunos = [];
  window._ml_tri = tri;
  window._ml_av  = nomeAv;

  document.getElementById('ml-titulo').textContent  = `${nomeAv} — ${_ml_turma.turma}`;
  document.getElementById('ml-av-info').textContent = `${nomeAv} · ${tri}º Tri`;
  document.getElementById('ml-valor').value         = '';
  document.getElementById('ml-aviso').style.display = 'none';
  document.getElementById('ml-tabela').classList.add('hidden');
  document.getElementById('ml-tbody').innerHTML     = '';
  document.getElementById('ml-btn-lancar').disabled = true;
  document.getElementById('modal-lancar').classList.remove('hidden');

  _mlCarregarDatas(tri, nomeAv);
}

function _mlFechar() {
  document.getElementById('modal-lancar').classList.add('hidden');
}

// ── Carregar datas e valor do RCO ─────────────────────────────────────────────

async function _mlCarregarDatas(tri, nomeAv) {
  const sel   = document.getElementById('ml-data');
  const aviso = document.getElementById('ml-aviso');
  sel.innerHTML = '<option value="">Carregando...</option>';
  sel.disabled  = true;

  try {
    const {escola, turma, disciplina} = _ml_turma;
    const [resDatas, resRco] = await Promise.all([
      window.pywebview.api.get_datas_aula(escola, turma, disciplina, tri),
      state.status?.api_conectada
        ? window.pywebview.api.get_avaliacoes_rco_api(escola, turma, disciplina)
        : Promise.resolve(null),
    ]);

    if (resDatas.ok && resDatas.datas.length > 0) {
      sel.innerHTML = '<option value="">Selecione a data...</option>' +
        resDatas.datas.map(d => `<option value="${d}">${d}</option>`).join('');
      sel.disabled = false;
    } else {
      sel.innerHTML = '<option value="">—</option>';
      if (!resDatas.ok) {
        aviso.textContent   = resDatas.erro || 'Erro ao carregar datas';
        aviso.style.color   = '#d97706';
        aviso.style.display = '';
      }
    }

    if (resRco && resRco.ok) {
      const av = resRco.avaliacoes.find(a => a.nome === nomeAv);
      if (av) {
        const campoValor = document.getElementById('ml-valor');
        if (av.peso && !campoValor.value) campoValor.value = av.peso;
        if (av.data) {
          for (const o of sel.options) {
            if (o.value === av.data) { sel.value = av.data; break; }
          }
        }
      }
    }
  } catch(e) {
    sel.innerHTML = '<option value="">—</option>';
  }
}

// ── Buscar notas da planilha ──────────────────────────────────────────────────

async function _mlBuscarPlanilha() {
  const tri    = window._ml_tri || 1;
  const coluna = window._ml_av  || '';
  if (!coluna) { toast('AV não definida', 'error'); return; }

  const btn = document.getElementById('ml-btn-buscar');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Buscando...';

  try {
    const {escola, turma, disciplina} = _ml_turma;
    const res = await window.pywebview.api.get_notas_planilha(escola, turma, disciplina, tri, coluna);
    if (!res.ok) { toast('Erro: ' + res.erro, 'error'); return; }

    _ml_alunos = res.alunos;
    _mlRenderTabela();
    document.getElementById('ml-tabela').classList.remove('hidden');
    document.getElementById('ml-btn-lancar').disabled = !state.chrome;
  } catch(e) {
    toast('Erro inesperado: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '📋 Buscar da planilha';
  }
}

// ── Renderizar tabela ─────────────────────────────────────────────────────────

function _mlRenderTabela() {
  document.getElementById('ml-tbody').innerHTML = _ml_alunos.map((a, i) => {
    const inBadge = a.inativo
      ? '<span style="font-size:10px;background:#fee2e2;color:#991b1b;padding:1px 6px;border-radius:10px;margin-left:6px">Inativo</span>'
      : '';
    const notaExibida = a.nota_rco && a.nota_rco !== '0'
      ? (parseFloat(a.nota_rco) / 10).toFixed(1).replace('.', ',')
      : '—';
    return `<tr style="${a.inativo ? 'opacity:.5' : ''}">
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;color:#94a3b8">${a.numero}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9">${a.nome}${inBadge}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center;color:#64748b">${notaExibida}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center">
        <input type="text" maxlength="3" ${a.inativo ? 'disabled' : ''}
               value="${a.nota_rco && a.nota_rco !== '0' ? a.nota_rco : ''}"
               oninput="_ml_alunos[${i}].nota_rco=this.value"
               style="width:60px;text-align:center;border:1px solid #e2e8f0;border-radius:6px;
                      padding:3px 6px;font-size:13px;background:${a.inativo?'#f8fafc':'#fff'}">
      </td>
    </tr>`;
  }).join('');
}

// ── Lançar notas ──────────────────────────────────────────────────────────────

async function _mlLancar() {
  if (!state.chrome) { toast('Conecte o Chrome primeiro', 'error'); return; }

  const tri    = window._ml_tri || 0;
  const tipoAv = window._ml_av  || '';
  if (!tri || !tipoAv) { toast('AV não definida', 'error'); return; }

  const data  = document.getElementById('ml-data').value;
  const valor = document.getElementById('ml-valor').value.trim();
  if (!data)                          { toast('Selecione a data da avaliação', 'error'); return; }
  if (!valor || isNaN(parseFloat(valor))) { toast('Informe o valor máximo', 'error'); return; }

  const notas = _ml_alunos
    .filter(a => !a.inativo && a.nota_rco && a.nota_rco !== '0')
    .map(a => ({
      nome_normalizado: a.nome.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim(),
      nota: a.nota_rco,
    }));

  if (notas.length === 0) { toast('Nenhuma nota para lançar', 'error'); return; }

  const triStr  = tri === 1 ? '1º Tri' : tri === 2 ? '2º Tri' : '3º Tri';
  const valorRco = String(Math.round(parseFloat(valor) * 10)).padStart(2, '0');

  const btn = document.getElementById('ml-btn-lancar');
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span>Lançando ${notas.length} notas...`;

  try {
    const {escola, turma, disciplina} = _ml_turma;
    const res = await window.pywebview.api.lancar_notas(
      escola, turma, disciplina, triStr, tipoAv, data, valorRco, notas
    );
    if (res.ok) {
      toast('Notas lançadas com sucesso!', 'success');
      _mlFechar();
    } else {
      toast('Erro: ' + res.erro, 'error');
    }
  } catch(e) {
    toast('Erro inesperado: ' + e, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Lançar no RCO';
  }
}
