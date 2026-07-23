/* ════════════════════════════════════════════════════════════════════
   PAINEL DE GESTÃO DE PROCESSOS — aplicação (dados + rotas + telas).
   Fonte de dados, em ordem de prioridade:
     1. Google Sheets (gviz JSONP), se PAINEL_CONFIG.googleSheetId
        estiver preenchido — mesmo padrão do Painel do PTD;
     2. Planilha local data/painel-processos-dados.xlsx (SheetJS);
     3. js/dados.js (window.PAINEL_DADOS), gerado por
        scripts/planilha_para_js.py — funciona até em file://.
   ════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  /* ── configuração ─────────────────────────────────────────────────── */
  var CONFIG = Object.assign({
    googleSheetId: '',                      // cole aqui o ID da planilha Google (opcional)
    arquivoXlsx: 'data/painel-processos-dados.xlsx',
    abas: ['Macroprocessos', 'Processos', 'Subprocessos', 'Atividades',
           'Documentos', 'Riscos', 'Indicadores', 'Diario_Mapeamento']
  }, window.PAINEL_CONFIG || {});

  var d = document;
  function $(s, c) { return (c || d).querySelector(s); }
  function $all(s, c) { return Array.prototype.slice.call((c || d).querySelectorAll(s)); }

  /* ── helpers ──────────────────────────────────────────────────────── */
  function esc(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function listar(v) {                       // "a; b; c" → ['a','b','c']
    if (v == null || v === '') return [];
    if (Array.isArray(v)) return v;
    return String(v).split(';').map(function (x) { return x.trim(); }).filter(Boolean);
  }
  function evidencias(v) {                   // "Nome|URL; Nome2|URL2"
    return listar(v).map(function (par) {
      var p = par.split('|');
      return { nome: p[0].trim(), url: (p[1] || '').trim() };
    });
  }
  function isoData(v) {                      // Date | 'YYYY-MM-DD' | 'Date(y,m,d)' | 'dd/mm/aaaa'
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') +
        '-' + String(v.getDate()).padStart(2, '0');
    }
    var s = String(v).trim();
    var g = s.match(/^Date\((\d+),(\d+),(\d+)/);
    if (g) return g[1] + '-' + String(+g[2] + 1).padStart(2, '0') + '-' + String(g[3]).padStart(2, '0');
    var br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (br) return br[3] + '-' + br[2].padStart(2, '0') + '-' + br[1].padStart(2, '0');
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return null;
  }
  function fmtData(iso) {
    if (!iso) return '—';
    var p = String(iso).slice(0, 10).split('-');
    return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(iso);
  }
  function hojeISO() { return isoData(new Date()); }
  function slug(s) {
    return String(s || '').toLowerCase().normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }
  function num(v) {
    if (v == null || v === '') return null;
    var n = typeof v === 'number' ? v : parseFloat(String(v).replace('%', '').replace(',', '.'));
    return isNaN(n) ? null : n;
  }
  function pctNorm(v) {                      // 0..1 ou 0..100 → 0..100
    var n = num(v); if (n == null) return 0;
    return Math.round(n <= 1 ? n * 100 : n);
  }
  function simNao(v) {
    if (v === true) return true;
    return /^s/i.test(String(v == null ? '' : v).trim());
  }

  /* ── camada de dados ──────────────────────────────────────────────── */
  var DADOS = null;   // coleções normalizadas
  var IDX = null;     // índices por código / vínculo
  var FONTE = '';     // descrição da origem carregada

  function gvizLinhas(table) {               // tabela gviz → [{Header: valor}]
    var cols = (table.cols || []).map(function (c) { return (c.label || c.id || '').trim(); });
    var rows = (table.rows || []).map(function (r) {
      return (r.c || []).map(function (c) { return c ? (c.v != null ? c.v : c.f) : null; });
    });
    // Se os labels vierem vazios, a 1ª linha é o cabeçalho
    if (!cols.some(Boolean) && rows.length) { cols = rows.shift().map(function (x) { return String(x || '').trim(); }); }
    return rows.filter(function (r) { return r[0] != null && String(r[0]).trim() !== ''; })
      .map(function (r) {
        var o = {};
        cols.forEach(function (k, i) { if (k) o[k] = r[i] != null ? r[i] : null; });
        return o;
      });
  }
  function carregarAbaGviz(aba) {            // JSONP — imune a CORS (padrão do painel do PTD)
    return new Promise(function (resolve, reject) {
      var cb = '_ppGviz' + Date.now() + Math.floor(Math.random() * 1e5);
      var done = false, sc = d.createElement('script');
      function fim() { try { delete window[cb]; } catch (e) {} if (sc.parentNode) sc.parentNode.removeChild(sc); }
      window[cb] = function (resp) {
        if (done) return; done = true; fim();
        if (resp && resp.status === 'ok') resolve(gvizLinhas(resp.table || {}));
        else reject(new Error('gviz: ' + ((resp && resp.errors && resp.errors[0] && resp.errors[0].detailed_message) || 'erro')));
      };
      sc.onerror = function () { if (!done) { done = true; fim(); reject(new Error('gviz: falha de rede')); } };
      setTimeout(function () { if (!done) { done = true; fim(); reject(new Error('gviz: tempo esgotado')); } }, 10000);
      sc.src = 'https://docs.google.com/spreadsheets/d/' + CONFIG.googleSheetId +
        '/gviz/tq?tqx=out:json;responseHandler:' + cb + '&sheet=' + encodeURIComponent(aba);
      d.head.appendChild(sc);
    });
  }
  function carregarGoogle() {
    return Promise.all(CONFIG.abas.map(carregarAbaGviz)).then(function (listas) {
      var o = {};
      CONFIG.abas.forEach(function (aba, i) { o[aba] = listas[i]; });
      FONTE = 'Google Sheets (tempo real)';
      return o;
    });
  }
  function carregarXlsx() {
    if (typeof XLSX === 'undefined') return Promise.reject(new Error('SheetJS indisponível'));
    return fetch(CONFIG.arquivoXlsx).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.arrayBuffer();
    }).then(function (buf) {
      var wb = XLSX.read(buf, { cellDates: true });
      var o = {};
      CONFIG.abas.forEach(function (aba) {
        var ws = wb.Sheets[aba];
        o[aba] = ws ? XLSX.utils.sheet_to_json(ws, { defval: null, raw: true })
          .filter(function (l) { var k = Object.keys(l)[0]; return l[k] != null && String(l[k]).trim() !== ''; }) : [];
      });
      FONTE = 'Planilha local (' + CONFIG.arquivoXlsx.split('/').pop() + ')';
      return o;
    });
  }
  function carregarEmbutido() {
    if (window.PAINEL_DADOS) {
      FONTE = 'Dados embutidos (js/dados.js' +
        (window.PAINEL_DADOS._gerado_em ? ', ' + window.PAINEL_DADOS._gerado_em : '') + ')';
      return Promise.resolve(window.PAINEL_DADOS);
    }
    return Promise.reject(new Error('js/dados.js ausente'));
  }
  function carregarDados() {
    var cadeia = CONFIG.googleSheetId
      ? carregarGoogle().catch(function (e) { console.warn('Google Sheets falhou:', e.message); return carregarXlsx(); })
      : carregarXlsx();
    return cadeia
      .catch(function (e) { console.warn('Planilha local falhou:', e.message); return carregarEmbutido(); })
      .then(normalizar);
  }

  /* ── normalização + índices ───────────────────────────────────────── */
  function normalizar(bruto) {
    function pega(aba) { return (bruto[aba] || []).map(function (l) { return Object.assign({}, l); }); }
    var dd = {
      macros: pega('Macroprocessos'),
      procs: pega('Processos'),
      subs: pega('Subprocessos'),
      ativs: pega('Atividades'),
      docs: pega('Documentos'),
      riscos: pega('Riscos'),
      inds: pega('Indicadores'),
      diario: pega('Diario_Mapeamento')
    };
    dd.macros.forEach(function (m) { m._cat = slug(m.Categoria); });
    dd.procs.forEach(function (p) {
      p.Percentual = pctNorm(p.Percentual);
      ['Inicio_Mapeamento', 'Prazo_Previsto', 'Data_Conclusao', 'Ultima_Atualizacao']
        .forEach(function (k) { p[k] = isoData(p[k]); });
      p._status = slug(p.Status_Mapeamento || 'Não iniciado');
    });
    dd.docs.forEach(function (x) { x.Data = isoData(x.Data); });
    dd.inds.forEach(function (x) {
      x.Ultima_Medicao = isoData(x.Ultima_Medicao);
      x.Meta = num(x.Meta); x.Resultado_Atual = num(x.Resultado_Atual);
      x._sit = situacaoInd(x);
    });
    dd.riscos.forEach(function (r) {
      r.Probabilidade_1a5 = num(r.Probabilidade_1a5) || 0;
      r.Impacto_1a5 = num(r.Impacto_1a5) || 0;
      r._nivel = r.Probabilidade_1a5 * r.Impacto_1a5;
      r._classe = classeRisco(r._nivel);
    });
    dd.diario.forEach(function (e) { e.Data = isoData(e.Data); });
    dd.diario.sort(function (a, b) { return String(b.Data || '').localeCompare(String(a.Data || '')); });

    var idx = { mp: {}, p: {}, sp: {}, a: {}, procsPorMacro: {}, subsPorProc: {}, ativsPorSub: {},
      vinc: { docs: {}, riscos: {}, inds: {} }, diarioPorProc: {} };
    dd.macros.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.macros.forEach(function (m) { idx.mp[m.Codigo] = m; });
    dd.procs.sort(function (a, b) { return String(a.Codigo).localeCompare(String(b.Codigo)); });
    dd.procs.forEach(function (p) {
      idx.p[p.Codigo] = p;
      (idx.procsPorMacro[p.Macroprocesso] = idx.procsPorMacro[p.Macroprocesso] || []).push(p);
    });
    dd.subs.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.subs.forEach(function (s) {
      idx.sp[s.Codigo] = s;
      (idx.subsPorProc[s.Processo] = idx.subsPorProc[s.Processo] || []).push(s);
    });
    dd.ativs.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.ativs.forEach(function (a) {
      idx.a[a.Codigo] = a;
      (idx.ativsPorSub[a.Subprocesso] = idx.ativsPorSub[a.Subprocesso] || []).push(a);
    });
    function vincula(mapa, item) {
      var ch = (item.Vinculo_Nivel || '') + '|' + (item.Vinculo_Codigo || '');
      (mapa[ch] = mapa[ch] || []).push(item);
    }
    dd.docs.forEach(function (x) { vincula(idx.vinc.docs, x); });
    dd.riscos.forEach(function (x) { vincula(idx.vinc.riscos, x); });
    dd.inds.forEach(function (x) { vincula(idx.vinc.inds, x); });
    dd.diario.forEach(function (e) {
      (idx.diarioPorProc[e.Processo] = idx.diarioPorProc[e.Processo] || []).push(e);
    });
    DADOS = dd; IDX = idx;
    return dd;
  }
  function vinculados(tipo, nivel, codigo) { return IDX.vinc[tipo][nivel + '|' + codigo] || []; }
  function classeRisco(n) {
    if (n >= 20) return 'Extremo';
    if (n >= 12) return 'Alto';
    if (n >= 5) return 'Moderado';
    return n > 0 ? 'Baixo' : '—';
  }
  function situacaoInd(x) {
    if (x.Resultado_Atual == null) return 'Sem medição';
    if (x.Meta == null) return 'Sem meta';
    var maior = /^maior/i.test(String(x.Polaridade || ''));
    var ok = maior ? x.Resultado_Atual >= x.Meta : x.Resultado_Atual <= x.Meta;
    return ok ? 'Meta atingida' : (maior ? 'Abaixo da meta' : 'Acima da meta');
  }

  /* ── componentes reutilizáveis (HTML) ─────────────────────────────── */
  function tagStatus(st) {
    var cls = slug(st || 'Não iniciado');
    return '<span class="tag-status ' + cls + '">' + esc(st || 'Não iniciado') + '</span>';
  }
  function tagCat(cat) {
    return '<span class="tag-cat ' + slug(cat) + '">' + esc(cat || '') + '</span>';
  }
  function tagNivel(cl) {
    return '<span class="nivel-tag nivel-' + slug(cl) + '">' + esc(cl) + '</span>';
  }
  function barraPct(p) {
    p = pctNorm(p);
    return '<div class="pct"><div class="trilho"><div class="barra" style="width:' + p +
      '%"></div></div><span class="valor">' + p + '%</span></div>';
  }
  function chips(str, icone) {
    var itens = listar(str);
    if (!itens.length) return '<span class="pp-vazio">Não informado</span>';
    return '<div class="chip-lista">' + itens.map(function (x) {
      return '<span class="chip">' + (icone ? '<i class="fas ' + icone + '" aria-hidden="true"></i> ' : '') + esc(x) + '</span>';
    }).join('') + '</div>';
  }
  function campo(rotulo, valorHtml, span2) {
    return '<div' + (span2 ? ' class="span2"' : '') + '><dt>' + esc(rotulo) + '</dt><dd>' +
      (valorHtml || '<span class="pp-vazio">Não informado</span>') + '</dd></div>';
  }
  function rotaDe(nivel, codigo) {
    var pre = { 'Macroprocesso': 'mp', 'Processo': 'p', 'Subprocesso': 'sp', 'Atividade': 'a' }[nivel];
    return pre ? '#/' + pre + '/' + encodeURIComponent(codigo) : '#/';
  }
  function nomeDe(nivel, codigo) {
    var it = nivel === 'Macroprocesso' ? IDX.mp[codigo] : nivel === 'Processo' ? IDX.p[codigo]
      : nivel === 'Subprocesso' ? IDX.sp[codigo] : IDX.a[codigo];
    return it ? it.Nome : codigo;
  }
  function linkVinculo(nivel, codigo) {
    return '<a href="' + rotaDe(nivel, codigo) + '"><span class="cod">' + esc(codigo) +
      '</span> ' + esc(nomeDe(nivel, codigo)) + '</a> <span class="pp-muted">(' + esc(nivel) + ')</span>';
  }
  function breadcrumb(trilha) {   // [{rotulo, href?}]
    return '<nav class="pp-breadcrumb" aria-label="Você está em">' +
      trilha.map(function (t, i) {
        var sep = i ? '<i class="fas fa-chevron-right" aria-hidden="true"></i>' : '';
        return sep + (t.href
          ? '<a href="' + t.href + '">' + esc(t.rotulo) + '</a>'
          : '<span class="atual" aria-current="page">' + esc(t.rotulo) + '</span>');
      }).join('') + '</nav>';
  }
  var MARCOS_ROTULOS = ['Formulário enviado', 'Formulário retornado', 'Reunião de contextualização',
    'AS-IS modelado', 'AS-IS validado', 'Normativos identificados', 'TO-BE elaborado',
    'TO-BE validado', 'Publicado no repositório'];
  var MARCOS_CAMPOS = ['M1_Formulario_Enviado', 'M2_Formulario_Retornado', 'M3_Reuniao_Contextualizacao',
    'M4_ASIS_Modelado', 'M5_ASIS_Validado', 'M6_Normativos_Identificados', 'M7_TOBE_Elaborado',
    'M8_TOBE_Validado', 'M9_Publicado_Repositorio'];
  function marcosHtml(p) {
    return '<ul class="marcos">' + MARCOS_CAMPOS.map(function (c, i) {
      var feito = simNao(p[c]);
      return '<li class="' + (feito ? 'feito' : '') + '"><span>' + esc(MARCOS_ROTULOS[i]) +
        '</span><i class="fas ' + (feito ? 'fa-check-circle' : 'fa-circle') + '" aria-hidden="true"></i></li>';
    }).join('') + '</ul>';
  }
  function diagramaHtml(caminho, titulo) {
    if (!caminho) return '<p class="pp-vazio">Diagrama BPMN ainda não publicado para este item.</p>';
    return '<figure class="diagrama-frame"><img src="' + esc(caminho) + '" alt="Diagrama BPMN (Bizagi) — ' +
      esc(titulo) + '" loading="lazy"></figure>' +
      '<div class="diagrama-acoes"><a class="br-button secondary small" href="' + esc(caminho) +
      '" target="_blank" rel="noopener"><i class="fas fa-up-right-and-down-left-from-center" aria-hidden="true"></i>&nbsp;Ampliar diagrama<span class="sr-only"> (abre em nova aba)</span></a></div>';
  }
  function listaDocsHtml(docs) {
    if (!docs.length) return '<p class="pp-vazio">Nenhum documento vinculado.</p>';
    var icones = { 'Diagrama BPMN': 'fa-diagram-project', 'Ata de reunião': 'fa-file-signature',
      'Relatório': 'fa-file-lines', 'Manual': 'fa-book', 'Norma interna': 'fa-scale-balanced',
      'Formulário/Modelo': 'fa-file-pen', 'Plano': 'fa-clipboard-list' };
    return docs.map(function (x) {
      var ic = icones[x.Tipo_Documento] || 'fa-file';
      var tit = x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) +
        '<span class="sr-only"> (abre em nova aba)</span></a>' : esc(x.Titulo);
      return '<div class="doc-item"><i class="fas ' + ic + ' fa-stack-ico" aria-hidden="true"></i><div>' +
        '<div class="tit">' + tit + '</div><div class="meta">' + esc(x.Tipo_Documento || 'Documento') +
        (x.Versao ? ' · v' + esc(x.Versao) : '') + (x.Data ? ' · ' + fmtData(x.Data) : '') +
        (x.Situacao ? ' · ' + esc(x.Situacao) : '') + '</div></div></div>';
    }).join('');
  }
  function tabelaRiscosHtml(riscos, comVinculo) {
    if (!riscos.length) return '<p class="pp-vazio">Nenhum risco registrado.</p>';
    return '<div class="pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th>' +
      (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Risco</th><th>P</th><th>I</th><th>P×I</th><th>Nível</th><th>Resposta</th><th>Status</th></tr></thead><tbody>' +
      riscos.map(function (r) {
        return '<tr id="risco-' + esc(r.ID) + '"><td class="cod">' + esc(r.ID) + '</td>' +
          (comVinculo ? '<td>' + linkVinculo(r.Vinculo_Nivel, r.Vinculo_Codigo) + '</td>' : '') +
          '<td>' + esc(r.Descricao_Risco) +
          (r.Controles_Tratamento ? '<div class="pp-muted" style="font-size:var(--fs-xs)">Tratamento: ' + esc(r.Controles_Tratamento) + '</div>' : '') +
          '</td><td>' + r.Probabilidade_1a5 + '</td><td>' + r.Impacto_1a5 + '</td><td><strong>' + r._nivel +
          '</strong></td><td>' + tagNivel(r._classe) + '</td><td>' + esc(r.Resposta || '—') +
          '</td><td>' + esc(r.Status || '—') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  function tabelaIndsHtml(inds, comVinculo) {
    if (!inds.length) return '<p class="pp-vazio">Nenhum indicador vinculado.</p>';
    return '<div class="pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th>' +
      (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Indicador</th><th>Meta</th><th>Resultado</th><th>Situação</th><th>Periodicidade</th><th>Última medição</th></tr></thead><tbody>' +
      inds.map(function (x) {
        var cls = x._sit === 'Meta atingida' ? 'sit-ok' : (x._sit === 'Sem medição' || x._sit === 'Sem meta') ? 'sit-neutra' : 'sit-ruim';
        var un = x.Unidade ? ' ' + esc(x.Unidade) : '';
        return '<tr><td class="cod">' + esc(x.ID) + '</td>' +
          (comVinculo ? '<td>' + linkVinculo(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' : '') +
          '<td><strong>' + esc(x.Nome) + '</strong>' +
          (x.Descricao_Formula ? '<div class="pp-muted" style="font-size:var(--fs-xs)">' + esc(x.Descricao_Formula) + '</div>' : '') +
          '</td><td>' + (x.Meta != null ? x.Meta + un : '—') + '</td><td>' + (x.Resultado_Atual != null ? x.Resultado_Atual + un : '—') +
          '</td><td class="' + cls + '">' + esc(x._sit) + '</td><td>' + esc(x.Periodicidade || '—') +
          '</td><td>' + fmtData(x.Ultima_Medicao) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  function timelineHtml(regs) {
    if (!regs.length) return '<p class="pp-vazio">Nenhum registro de mapeamento ainda.</p>';
    return '<ol class="timeline">' + regs.map(function (e) {
      var evs = evidencias(e.Evidencias);
      return '<li><div class="quando"><span class="tipo">' + esc(e.Tipo || 'Nota') + '</span> ' +
        fmtData(e.Data) + (e.Autor ? ' · ' + esc(e.Autor) : '') +
        (e.Processo ? ' · <a href="#/p/' + encodeURIComponent(e.Processo) + '">' + esc(e.Processo) + '</a>' : '') +
        '</div><h4>' + esc(e.Titulo || '(sem título)') + '</h4>' +
        (e.Descricao ? '<p style="font-size:var(--fs-sm)">' + esc(e.Descricao) + '</p>' : '') +
        (e.Participantes ? '<p class="pp-muted" style="font-size:var(--fs-xs)"><i class="fas fa-users" aria-hidden="true"></i> ' + esc(e.Participantes) + '</p>' : '') +
        ((e.Entradas_Insumos || e.Saidas_Entregaveis)
          ? '<div class="es-grid">' +
            '<div class="es-caixa"><b><i class="fas fa-arrow-right-to-bracket" aria-hidden="true"></i> Entradas / insumos</b>' + (listar(e.Entradas_Insumos).map(esc).join('; ') || '—') + '</div>' +
            '<div class="es-caixa"><b><i class="fas fa-arrow-right-from-bracket" aria-hidden="true"></i> Saídas / entregáveis</b>' + (listar(e.Saidas_Entregaveis).map(esc).join('; ') || '—') + '</div></div>'
          : '') +
        (evs.length ? '<p class="evidencias" style="margin-top:6px"><b style="font-size:10px;text-transform:uppercase;color:var(--gray-60)">Evidências:</b> ' +
          evs.map(function (ev) {
            return ev.url ? '<a href="' + esc(ev.url) + '" target="_blank" rel="noopener"><i class="fas fa-paperclip" aria-hidden="true"></i> ' + esc(ev.nome) + '<span class="sr-only"> (abre em nova aba)</span></a>'
              : '<span class="chip"><i class="fas fa-paperclip" aria-hidden="true"></i> ' + esc(ev.nome) + '</span>';
          }).join(' ') + '</p>' : '') +
        (e.Memoria ? '<p class="memoria"><i class="fas fa-book-open" aria-hidden="true"></i> ' + esc(e.Memoria) + '</p>' : '') +
        '</li>';
    }).join('') + '</ol>';
  }

  /* ── roteador + abas ──────────────────────────────────────────────── */
  var ROTAS_ABA = { inicio: '#/', catalogo: '#/catalogo', documentos: '#/documentos',
    riscos: '#/riscos', indicadores: '#/indicadores', diario: '#/diario', metodologia: '#/metodologia' };
  function mostrarPainel(id) {
    $all('#mainTabContent > .tab-panel').forEach(function (p) {
      var ativo = p.id === 'panel-' + id;
      p.classList.toggle('active', ativo);
      p.hidden = !ativo;
    });
    $all('.tab-nav [data-rota]').forEach(function (b) {
      var ativo = b.getAttribute('data-painel') === id ||
        (id === 'detalhe' && b.getAttribute('data-painel') === 'catalogo') ||
        (id === 'busca' && b.getAttribute('data-painel') === 'inicio');
      b.setAttribute('aria-selected', ativo ? 'true' : 'false');
      b.closest('.tab-item').classList.toggle('active', ativo);
    });
  }
  function rota() {
    var h = location.hash || '#/';
    var m;
    if (h === '#' || h === '#/') { renderInicio(); mostrarPainel('inicio'); }
    else if (h === '#/catalogo') { renderCatalogo(); mostrarPainel('catalogo'); }
    else if (h === '#/documentos') { renderDocumentos(); mostrarPainel('documentos'); }
    else if (h === '#/riscos') { renderRiscos(); mostrarPainel('riscos'); }
    else if (h === '#/indicadores') { renderIndicadores(); mostrarPainel('indicadores'); }
    else if (h === '#/diario') { renderDiario(); mostrarPainel('diario'); }
    else if (h === '#/metodologia') { renderMetodologia(); mostrarPainel('metodologia'); }
    else if ((m = h.match(/^#\/busca\?q=(.*)$/))) { renderBusca(decodeURIComponent(m[1])); mostrarPainel('busca'); }
    else if ((m = h.match(/^#\/(mp|p|sp|a)\/(.+)$/))) { renderDetalhe(m[1], decodeURIComponent(m[2])); mostrarPainel('detalhe'); }
    else { renderInicio(); mostrarPainel('inicio'); }
    var alvo = $('#navigation');
    if (alvo && window.scrollY > alvo.offsetTop) alvo.scrollIntoView();
  }
  d.addEventListener('click', function (ev) {
    var b = ev.target.closest('.tab-nav [data-rota]');
    if (b) location.hash = b.getAttribute('data-rota');
  });

  /* ── alertas operacionais (sino) ──────────────────────────────────── */
  window.showNotifTab = function (qual, btn) {
    $('#notifPanelOverdue').hidden = qual !== 'overdue';
    $('#notifPanelOverdue').classList.toggle('active', qual === 'overdue');
    $('#notifPanelRisk').hidden = qual !== 'risk';
    $('#notifPanelRisk').classList.toggle('active', qual === 'risk');
    $all('#notifTabs .tab-item').forEach(function (li) { li.classList.remove('active'); });
    if (btn) { btn.closest('.tab-item').classList.add('active'); }
  };
  function montarAlertas() {
    var hoje = hojeISO();
    var vencidos = DADOS.procs.filter(function (p) {
      return p.Prazo_Previsto && p.Prazo_Previsto < hoje && p._status !== 'concluido';
    });
    var criticos = DADOS.riscos.filter(function (r) {
      return (r._classe === 'Alto' || r._classe === 'Extremo') && !/encerrad/i.test(String(r.Status || ''));
    });
    $('#notifOverdueCount').textContent = vencidos.length;
    $('#notifRiskCount').textContent = criticos.length;
    var badge = $('#notifBadge');
    var total = vencidos.length + criticos.length;
    badge.textContent = total; badge.hidden = !total;
    $('#notifOverdueList').innerHTML = vencidos.length ? vencidos.map(function (p) {
      return '<a class="br-item" href="#/p/' + encodeURIComponent(p.Codigo) + '"><strong>' + esc(p.Codigo) +
        '</strong> — ' + esc(p.Nome) + '<br><span class="pp-muted">Prazo: ' + fmtData(p.Prazo_Previsto) + '</span></a>';
    }).join('') : '<div class="br-item pp-vazio">Nenhum prazo vencido.</div>';
    $('#notifRiskList').innerHTML = criticos.length ? criticos.map(function (r) {
      return '<a class="br-item" href="' + rotaDe(r.Vinculo_Nivel, r.Vinculo_Codigo) + '"><strong>' + esc(r.ID) +
        '</strong> ' + tagNivel(r._classe) + '<br>' + esc(r.Descricao_Risco) + '</a>';
    }).join('') : '<div class="br-item pp-vazio">Nenhum risco crítico aberto.</div>';
    var btnFechar = $('#notifCloseBtn');
    if (btnFechar) btnFechar.onclick = function () { $('#notifPanel').hidden = true; };
  }

  /* ── ações do cabeçalho ───────────────────────────────────────────── */
  window.refreshAll = function () {
    var chip = $('#syncChip'); if (chip) chip.textContent = 'Atualizando…';
    carregarDados().then(function () { posCarga(); }).catch(function (e) {
      if (chip) chip.textContent = 'Falha ao atualizar';
      console.error(e);
    });
  };
  function ligarAcoesCabecalho() {
    var exp = $('#actExportCsv');
    if (exp) exp.onclick = function () {
      var cab = ['Codigo', 'Macroprocesso', 'Nome', 'Status_Mapeamento', 'Percentual', 'Fase_Ciclo_BPM',
        'Area_Responsavel', 'Dono_Processo', 'Prazo_Previsto'];
      var linhas = [cab.join(';')].concat(DADOS.procs.map(function (p) {
        return cab.map(function (k) { return '"' + String(p[k] == null ? '' : p[k]).replace(/"/g, '""') + '"'; }).join(';');
      }));
      var blob = new Blob(['\ufeff' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8' });
      var a = d.createElement('a');
      a.href = URL.createObjectURL(blob); a.download = 'processos-painel.csv';
      d.body.appendChild(a); a.click(); a.remove();
    };
    var cop = $('#actCopyLink');
    if (cop) cop.onclick = function () {
      navigator.clipboard && navigator.clipboard.writeText(location.href.split('#')[0]);
      cop.innerHTML = '<i class="fas fa-check" aria-hidden="true"></i> Link copiado!';
      setTimeout(function () { cop.innerHTML = '<i class="fas fa-link" aria-hidden="true"></i> Copiar link do painel'; }, 2000);
    };
    var abrir = $('#actOpenSheet');
    if (abrir) abrir.href = CONFIG.googleSheetId
      ? 'https://docs.google.com/spreadsheets/d/' + CONFIG.googleSheetId + '/edit'
      : CONFIG.arquivoXlsx;
    var busca = $('#headerSearchInput'), envia = $('#headerSearchSubmit');
    function irBusca() {
      var q = (busca.value || '').trim();
      if (q) { location.hash = '#/busca?q=' + encodeURIComponent(q); $('#headerSearch').classList.remove('active'); }
    }
    if (busca) busca.addEventListener('keydown', function (e) { if (e.key === 'Enter') irBusca(); });
    if (envia) envia.onclick = irBusca;
  }

  /* ── TELA: início (KPIs + cadeia de valor) ────────────────────────── */
  var INSTITUCIONAL = {
    missao: 'Promover o desenvolvimento regional de forma integrada e sustentável nas bacias hidrográficas, contribuindo para a redução das desigualdades.',
    visao: 'Ser referência na execução de políticas públicas para o desenvolvimento regional.',
    proposito: 'Construir caminhos de oportunidades, integrando regiões e transformando vidas de forma planejada e sustentável.',
    valores: ['Foco na sociedade', 'Excelência', 'Transparência', 'Valorização dos Colaboradores',
      'Sustentabilidade', 'Ética', 'Comprometimento', 'Estímulo à Diversidade', 'Inovação']
  };
  function blocoCadeia(titulo, cor, itens) {
    return '<div class="cv-bloco"><div class="cv-titulo" style="background:' + cor + '">' + titulo + '</div><ul>' +
      itens.map(function (m) {
        return '<li><a href="#/mp/' + encodeURIComponent(m.Codigo) + '"><span class="cod">' + esc(m.Codigo) +
          '</span><div class="nome">' + esc(m.Nome) + '</div></a></li>';
      }).join('') + '</ul></div>';
  }
  function renderInicio() {
    var el = $('#viewInicio');
    var procs = DADOS.procs;
    var concl = procs.filter(function (p) { return p._status === 'concluido'; }).length;
    var andamento = procs.filter(function (p) { return p._status === 'em-andamento'; }).length;
    var media = procs.length ? Math.round(procs.reduce(function (s, p) { return s + p.Percentual; }, 0) / procs.length) : 0;
    var criticos = DADOS.riscos.filter(function (r) {
      return (r._classe === 'Alto' || r._classe === 'Extremo') && !/encerrad/i.test(String(r.Status || ''));
    }).length;
    var docsVig = DADOS.docs.filter(function (x) { return /vigente/i.test(String(x.Situacao || '')); }).length;
    var ger = DADOS.macros.filter(function (m) { return m._cat === 'gerencial'; });
    var fin = DADOS.macros.filter(function (m) { return m._cat === 'finalistico'; });
    var sup = DADOS.macros.filter(function (m) { return m._cat === 'suporte'; });
    el.innerHTML =
      '<section class="pp-hero"><span class="eyebrow">Gestão de processos · Gestão por processos · BPM CBOK 4.0 · PMBOK</span>' +
      '<h1>Cadeia de valor e mapeamento de processos da Codevasf</h1>' +
      '<p>Consulte a hierarquia completa — do macroprocesso à atividade — com fichas, diagramas BPMN (Bizagi), ' +
      'documentos, riscos, indicadores e o registro rastreável de cada mapeamento realizado.</p>' +
      '<div class="acoes"><a class="br-button inverted" href="#/catalogo"><i class="fas fa-layer-group" aria-hidden="true"></i>&nbsp;Catálogo de processos</a>' +
      '<a class="br-button outline-inv" href="#/diario"><i class="fas fa-timeline" aria-hidden="true"></i>&nbsp;Diário de mapeamento</a>' +
      '<a class="br-button outline-inv" href="#/metodologia"><i class="fas fa-book" aria-hidden="true"></i>&nbsp;Metodologia</a></div></section>' +
      '<div class="kpi-grid">' +
      '<div class="kpi"><span class="num">' + DADOS.macros.length + '</span><span class="lbl">Macroprocessos</span><span class="sub">' + procs.length + ' processos · ' + DADOS.subs.length + ' subprocessos · ' + DADOS.ativs.length + ' atividades</span></div>' +
      '<div class="kpi ok"><span class="num">' + concl + '</span><span class="lbl">Mapeamentos concluídos</span><span class="sub">' + andamento + ' em andamento</span></div>' +
      '<div class="kpi"><span class="num">' + media + '%</span><span class="lbl">Avanço médio</span><span class="sub">do mapeamento da carteira</span></div>' +
      '<div class="kpi ' + (criticos ? 'erro' : 'ok') + '"><span class="num">' + criticos + '</span><span class="lbl">Riscos críticos abertos</span><span class="sub">nível Alto ou Extremo</span></div>' +
      '<div class="kpi"><span class="num">' + docsVig + '</span><span class="lbl">Documentos vigentes</span><span class="sub">' + DADOS.diario.length + ' registros no diário</span></div>' +
      '</div>' +
      '<section class="pp-sec" id="sec-cadeia"><div class="pp-sec-h"><h2>Cadeia de Valor Integrada</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="cadeia">' +
      '<aside class="cv-aside cv-missao"><h3>Missão</h3><p>' + esc(INSTITUCIONAL.missao) + '</p><h3>Visão</h3><p>' + esc(INSTITUCIONAL.visao) + '</p></aside>' +
      '<div class="cv-centro">' + blocoCadeia('Macroprocessos Gerenciais', 'var(--pp-gerencial)', ger) +
      blocoCadeia('Macroprocessos Finalísticos', 'var(--pp-finalistico)', fin) + '</div>' +
      '<aside class="cv-aside cv-proposito"><h3>Propósito</h3><p>' + esc(INSTITUCIONAL.proposito) + '</p></aside>' +
      '<div class="cv-suporte">' + blocoCadeia('Macroprocessos de Suporte', 'var(--pp-suporte)', sup) + '</div>' +
      '<div class="cv-valores"><strong>Valores</strong>' + INSTITUCIONAL.valores.map(function (v) { return '<span>· ' + esc(v) + '</span>'; }).join('') + '</div>' +
      '</div><p class="pp-muted" style="margin-top:var(--sp2);font-size:var(--fs-xs)">Clique em um macroprocesso para abrir a ficha e navegar até processos, subprocessos e atividades. Classificação conforme o BPM CBOK 4.0 (processos gerenciais, primários/finalísticos e de suporte).</p></section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Últimos registros do mapeamento</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card">' + timelineHtml(DADOS.diario.slice(0, 3)) +
      '<p style="margin-top:var(--sp2)"><a class="br-button secondary small" href="#/diario">Ver diário completo</a></p></div></section>';
  }

  /* ── TELA: catálogo ───────────────────────────────────────────────── */
  var filtroCat = { macro: '', status: '', q: '' };
  function cardProcesso(p) {
    return '<a class="proc-card" href="#/p/' + encodeURIComponent(p.Codigo) + '">' +
      '<div class="topo"><div><span class="cod" style="font-family:var(--noto-mono,monospace);font-size:var(--fs-xs);color:var(--gray-60)">' + esc(p.Codigo) + '</span>' +
      '<div class="nome">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
      '<div class="pp-muted" style="font-size:var(--fs-xs);margin-top:4px">' + esc(p.Area_Responsavel || '') +
      (p.Fase_Ciclo_BPM ? ' · ' + esc(p.Fase_Ciclo_BPM) : '') + '</div>' +
      '<div class="rodape">' + barraPct(p.Percentual) + '</div></a>';
  }
  function renderCatalogo() {
    var el = $('#viewCatalogo');
    var lista = DADOS.procs.filter(function (p) {
      if (filtroCat.macro && p.Macroprocesso !== filtroCat.macro) return false;
      if (filtroCat.status && slug(p.Status_Mapeamento) !== filtroCat.status) return false;
      if (filtroCat.q) {
        var q = filtroCat.q.toLowerCase();
        if ((p.Codigo + ' ' + p.Nome + ' ' + (p.Descricao || '')).toLowerCase().indexOf(q) < 0) return false;
      }
      return true;
    });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Catálogo de processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-filtros" role="search">' +
      '<label class="sr-only" for="fMacro">Filtrar por macroprocesso</label>' +
      '<select id="fMacro"><option value="">Todos os macroprocessos</option>' +
      DADOS.macros.map(function (m) {
        return '<option value="' + esc(m.Codigo) + '"' + (filtroCat.macro === m.Codigo ? ' selected' : '') + '>' + esc(m.Codigo + ' — ' + m.Nome) + '</option>';
      }).join('') + '</select>' +
      '<label class="sr-only" for="fStatus">Filtrar por status</label>' +
      '<select id="fStatus"><option value="">Todos os status</option>' +
      ['Não iniciado', 'Em andamento', 'Concluído', 'Suspenso'].map(function (s) {
        return '<option value="' + slug(s) + '"' + (filtroCat.status === slug(s) ? ' selected' : '') + '>' + s + '</option>';
      }).join('') + '</select>' +
      '<label class="sr-only" for="fBusca">Buscar no catálogo</label>' +
      '<input type="search" id="fBusca" placeholder="Buscar por código ou nome…" value="' + esc(filtroCat.q) + '">' +
      '<span class="pp-muted" style="font-size:var(--fs-sm)">' + lista.length + ' de ' + DADOS.procs.length + ' processos</span></div>' +
      (lista.length ? '<div class="proc-grid">' + lista.map(cardProcesso).join('') + '</div>'
        : '<p class="pp-vazio">Nenhum processo corresponde aos filtros. Limpe os filtros para ver todos.</p>');
    $('#fMacro').onchange = function () { filtroCat.macro = this.value; renderCatalogo(); };
    $('#fStatus').onchange = function () { filtroCat.status = this.value; renderCatalogo(); };
    $('#fBusca').oninput = function () { filtroCat.q = this.value; renderCatalogo(); $('#fBusca').focus(); };
  }

  /* ── TELA: detalhe (mp | p | sp | a) ──────────────────────────────── */
  function secVinculos(nivel, codigo) {
    var docs = vinculados('docs', nivel, codigo);
    var riscos = vinculados('riscos', nivel, codigo);
    var inds = vinculados('inds', nivel, codigo);
    return '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> Indicadores de desempenho</h3>' + tabelaIndsHtml(inds, false) + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Riscos (matriz 5×5 · P×I)</h3>' + tabelaRiscosHtml(riscos, false) + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-folder-open" aria-hidden="true"></i> Documentos vinculados</h3>' + listaDocsHtml(docs) + '</div>';
  }
  function renderDetalhe(tipo, cod) {
    var el = $('#viewDetalhe');
    var CORES = { gerencial: 'var(--pp-gerencial)', finalistico: 'var(--pp-finalistico)', suporte: 'var(--pp-suporte)' };
    if (tipo === 'mp') {
      var m = IDX.mp[cod];
      if (!m) { el.innerHTML = naoEncontrado('Macroprocesso', cod); return; }
      var filhos = IDX.procsPorMacro[cod] || [];
      var media = filhos.length ? Math.round(filhos.reduce(function (s, p) { return s + p.Percentual; }, 0) / filhos.length) : 0;
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }, { rotulo: m.Codigo + ' — ' + m.Nome }]) +
        '<section class="ficha-hero" style="background:' + (CORES[m._cat] || 'var(--cv-navy)') + '">' +
        '<span class="eyebrow">Macroprocesso ' + esc(m.Categoria) + '</span><h2>' + esc(m.Codigo) + ' — ' + esc(m.Nome) + '</h2>' +
        '<div class="meta">' + tagCat(m.Categoria) + '<span>' + filhos.length + ' processos vinculados</span><span>· mapeamento médio ' + media + '%</span></div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do macroprocesso</h3><dl class="ficha-dl">' +
        campo('Objetivo', m.Objetivo && esc(m.Objetivo)) + campo('Descrição', m.Descricao && esc(m.Descricao)) +
        campo('Unidade responsável', m.Unidade_Responsavel && esc(m.Unidade_Responsavel)) +
        campo('Dono do processo (process owner)', m.Dono_Processo && esc(m.Dono_Processo)) +
        campo('Entregas (produtos/serviços)', chips(m.Entregas), true) +
        campo('Clientes / beneficiários', chips(m.Clientes_Beneficiarios)) +
        campo('Partes interessadas', chips(m.Partes_Interessadas)) +
        campo('Sistemas utilizados', chips(m.Sistemas, 'fa-desktop')) +
        campo('Normativos aplicáveis', chips(m.Normativos_Aplicaveis, 'fa-scale-balanced'), true) +
        (m.Observacoes ? campo('Observações', esc(m.Observacoes), true) : '') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(m.Imagem_Bizagi, m.Nome) + '</div>' +
        secVinculos('Macroprocesso', cod) +
        '</div><aside>' +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Processos vinculados</h3>' +
        (filhos.length ? filhos.map(function (p) {
          return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="#/p/' + encodeURIComponent(p.Codigo) + '">' +
            '<div class="topo"><div><span style="font-family:var(--noto-mono,monospace);font-size:var(--fs-xs);color:var(--gray-60)">' + esc(p.Codigo) + '</span>' +
            '<div class="nome" style="font-size:var(--fs-sm)">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
            '<div class="rodape">' + barraPct(p.Percentual) + '</div></a>';
        }).join('') : '<p class="pp-vazio">Nenhum processo cadastrado.</p>') + '</div></aside></div>';
      return;
    }
    if (tipo === 'p') {
      var p = IDX.p[cod];
      if (!p) { el.innerHTML = naoEncontrado('Processo', cod); return; }
      var mp = IDX.mp[p.Macroprocesso];
      var subs = IDX.subsPorProc[cod] || [];
      var regs = IDX.diarioPorProc[cod] || [];
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
          .concat(mp ? [{ rotulo: mp.Codigo, href: '#/mp/' + encodeURIComponent(mp.Codigo) }] : [])
          .concat([{ rotulo: p.Codigo + ' — ' + p.Nome }])) +
        '<section class="ficha-hero" style="background:var(--cv-navy)">' +
        '<span class="eyebrow">Processo' + (mp ? ' · ' + esc(mp.Categoria) : '') + '</span>' +
        '<h2>' + esc(p.Codigo) + ' — ' + esc(p.Nome) + '</h2>' +
        '<div class="meta">' + tagStatus(p.Status_Mapeamento) +
        '<span>Mapeamento: <strong>' + p.Percentual + '%</strong></span>' +
        (p.Fase_Ciclo_BPM ? '<span class="chip" style="background:rgba(255,255,255,.15);border-color:rgba(255,255,255,.4);color:#fff">' + esc(p.Fase_Ciclo_BPM) + '</span>' : '') +
        (p.Area_Responsavel ? '<span>Área: ' + esc(p.Area_Responsavel) + '</span>' : '') +
        (p.Processo_SEI ? '<span><i class="fas fa-file-lines" aria-hidden="true"></i> SEI ' + esc(p.Processo_SEI) + '</span>' : '') +
        '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-bullseye" aria-hidden="true"></i> Descrição e objetivo</h3>' +
        '<p style="font-size:var(--fs-sm)">' + esc(p.Descricao || '') + '</p>' +
        (p.Objetivo ? '<p style="font-size:var(--fs-sm);margin-top:var(--sp1)"><strong>Objetivo:</strong> ' + esc(p.Objetivo) + '</p>' : '') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-right-left" aria-hidden="true"></i> SIPOC</h3><div class="sipoc">' +
        '<div class="col"><h4>Fornecedores</h4><ul>' + (listar(p.Fornecedores).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Entradas</h4><ul>' + (listar(p.Entradas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col centro"><h4>Processo</h4><div style="font-weight:600">' + esc(p.Nome) + '</div></div>' +
        '<div class="col"><h4>Saídas</h4><ul>' + (listar(p.Saidas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Clientes</h4><ul>' + (listar(p.Clientes).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '</div></div>' +
        '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M9)</h3>' + marcosHtml(p) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(p.Imagem_Bizagi, p.Nome) + '</div>' +
        secVinculos('Processo', cod) +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos</h3>' +
        (subs.length ? '<div class="pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Subprocesso</th><th>Entregas</th><th></th></tr></thead><tbody>' +
          subs.map(function (s) {
            return '<tr><td class="cod">' + esc(s.Codigo) + '</td><td><a href="#/sp/' + encodeURIComponent(s.Codigo) + '"><strong>' + esc(s.Nome) + '</strong></a>' +
              (s.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-xs)">' + esc(s.Descricao) + '</div>' : '') + '</td>' +
              '<td style="font-size:var(--fs-xs)">' + (listar(s.Entregas).map(esc).join('; ') || '—') + '</td>' +
              '<td><a class="br-button secondary small" href="#/sp/' + encodeURIComponent(s.Codigo) + '">Abrir ficha</a></td></tr>';
          }).join('') + '</tbody></table></div>' : '<p class="pp-vazio">Nenhum subprocesso cadastrado.</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-timeline" aria-hidden="true"></i> Diário de mapeamento deste processo</h3>' + timelineHtml(regs) + '</div>' +
        '</div><aside>' +
        '<div class="pp-card"><h3><i class="fas fa-calendar-check" aria-hidden="true"></i> Cronograma do projeto</h3>' + barraPct(p.Percentual) +
        '<dl class="ficha-dl" style="grid-template-columns:1fr 1fr;margin-top:var(--sp2)">' +
        campo('Início', fmtData(p.Inicio_Mapeamento)) + campo('Prazo', fmtData(p.Prazo_Previsto)) +
        campo('Conclusão', fmtData(p.Data_Conclusao)) + campo('Atualizado em', fmtData(p.Ultima_Atualizacao)) + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-users" aria-hidden="true"></i> Responsáveis</h3><dl class="ficha-dl" style="grid-template-columns:1fr">' +
        campo('Dono do processo', p.Dono_Processo && esc(p.Dono_Processo)) +
        campo('Área responsável', p.Area_Responsavel && esc(p.Area_Responsavel)) +
        campo('Interlocutor do mapeamento', p.Interlocutor && esc(p.Interlocutor)) +
        campo('Prioridade / complexidade', esc(p.Prioridade || '—') + ' / ' + esc(p.Complexidade || '—')) + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-forward" aria-hidden="true"></i> Próxima ação</h3>' +
        (p.Proxima_Acao ? '<p style="font-size:var(--fs-sm)">' + esc(p.Proxima_Acao) + '</p>' : '<p class="pp-vazio">—</p>') +
        (p.Pendencia ? '<div class="pp-aviso" style="margin:var(--sp2) 0 0"><strong>Pendência:</strong> ' + esc(p.Pendencia) + '</div>' : '') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-scale-balanced" aria-hidden="true"></i> Normativos relacionados</h3>' + chips(p.Normativos_Relacionados, 'fa-scale-balanced') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-desktop" aria-hidden="true"></i> Sistemas</h3>' + chips(p.Sistemas, 'fa-desktop') + '</div>' +
        '</aside></div>';
      return;
    }
    if (tipo === 'sp') {
      var s = IDX.sp[cod];
      if (!s) { el.innerHTML = naoEncontrado('Subprocesso', cod); return; }
      var pp = IDX.p[s.Processo]; var mpp = pp && IDX.mp[pp.Macroprocesso];
      var ativs = IDX.ativsPorSub[cod] || [];
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
          .concat(mpp ? [{ rotulo: mpp.Codigo, href: '#/mp/' + encodeURIComponent(mpp.Codigo) }] : [])
          .concat(pp ? [{ rotulo: pp.Codigo, href: '#/p/' + encodeURIComponent(pp.Codigo) }] : [])
          .concat([{ rotulo: s.Codigo + ' — ' + s.Nome }])) +
        '<section class="ficha-hero" style="background:var(--cv-blue)">' +
        '<span class="eyebrow">Subprocesso' + (pp ? ' de ' + esc(pp.Nome) : '') + '</span>' +
        '<h2>' + esc(s.Codigo) + ' — ' + esc(s.Nome) + '</h2>' +
        '<div class="meta"><span>' + ativs.length + ' atividades mapeadas</span>' +
        (s.Unidade_Responsavel ? '<span>· ' + esc(s.Unidade_Responsavel) + '</span>' : '') + '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do subprocesso</h3><dl class="ficha-dl">' +
        campo('Descrição', s.Descricao && esc(s.Descricao), true) +
        campo('Objetivo', s.Objetivo && esc(s.Objetivo), true) +
        campo('Unidade responsável', s.Unidade_Responsavel && esc(s.Unidade_Responsavel)) +
        campo('Dono', s.Dono && esc(s.Dono)) +
        campo('Entregas', chips(s.Entregas)) + campo('Sistemas', chips(s.Sistemas, 'fa-desktop')) + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Atividades (com entradas e saídas)</h3>' +
        (ativs.length ? '<div class="pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>#</th><th>Atividade</th><th>Responsável (ator)</th><th>Entradas</th><th>Saídas</th><th>Prazo</th></tr></thead><tbody>' +
          ativs.map(function (a, i) {
            return '<tr data-link><td>' + (i + 1) + '</td><td><a href="#/a/' + encodeURIComponent(a.Codigo) + '"><strong>' + esc(a.Nome) + '</strong></a>' +
              '<div class="cod">' + esc(a.Codigo) + '</div></td><td style="font-size:var(--fs-xs)">' + esc(a.Responsavel_Ator || '—') + '</td>' +
              '<td style="font-size:var(--fs-xs)">' + (listar(a.Entradas).map(esc).join('; ') || '—') + '</td>' +
              '<td style="font-size:var(--fs-xs)">' + (listar(a.Saidas).map(esc).join('; ') || '—') + '</td>' +
              '<td style="font-size:var(--fs-xs);white-space:nowrap">' + esc(a.Prazo_Padrao || '—') + '</td></tr>';
          }).join('') + '</tbody></table></div>' : '<p class="pp-vazio">Nenhuma atividade cadastrada.</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(s.Imagem_Bizagi, s.Nome) + '</div>' +
        secVinculos('Subprocesso', cod) +
        '</div><aside>' +
        (pp ? '<div class="pp-card"><h3><i class="fas fa-arrow-turn-up" aria-hidden="true"></i> Processo pai</h3>' +
          '<a class="proc-card" href="#/p/' + encodeURIComponent(pp.Codigo) + '"><div class="topo"><div><span class="cod">' + esc(pp.Codigo) + '</span><div class="nome" style="font-size:var(--fs-sm)">' + esc(pp.Nome) + '</div></div>' + tagStatus(pp.Status_Mapeamento) + '</div></a></div>' : '') +
        '</aside></div>';
      // clique na linha abre a atividade
      $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
        tr.addEventListener('click', function (ev) {
          if (ev.target.closest('a')) return;
          var a = tr.querySelector('a'); if (a) location.hash = a.getAttribute('href');
        });
      });
      return;
    }
    // tipo === 'a'
    var a = IDX.a[cod];
    if (!a) { el.innerHTML = naoEncontrado('Atividade', cod); return; }
    var sp2 = IDX.sp[a.Subprocesso]; var p2 = sp2 && IDX.p[sp2.Processo]; var mp2 = p2 && IDX.mp[p2.Macroprocesso];
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp2 ? [{ rotulo: mp2.Codigo, href: '#/mp/' + encodeURIComponent(mp2.Codigo) }] : [])
        .concat(p2 ? [{ rotulo: p2.Codigo, href: '#/p/' + encodeURIComponent(p2.Codigo) }] : [])
        .concat(sp2 ? [{ rotulo: sp2.Codigo, href: '#/sp/' + encodeURIComponent(sp2.Codigo) }] : [])
        .concat([{ rotulo: a.Codigo }])) +
      '<section class="ficha-hero" style="background:var(--pp-verde)">' +
      '<span class="eyebrow">Atividade' + (sp2 ? ' do subprocesso ' + esc(sp2.Nome) : '') + '</span>' +
      '<h2>' + esc(a.Codigo) + ' — ' + esc(a.Nome) + '</h2>' +
      '<div class="meta">' + (a.Responsavel_Ator ? '<span><i class="fas fa-user" aria-hidden="true"></i> ' + esc(a.Responsavel_Ator) + '</span>' : '') +
      (a.Prazo_Padrao ? '<span>· Prazo padrão: ' + esc(a.Prazo_Padrao) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da atividade</h3><dl class="ficha-dl">' +
      campo('Descrição', a.Descricao && esc(a.Descricao), true) +
      campo('Entradas (insumos)', chips(a.Entradas, 'fa-arrow-right-to-bracket')) +
      campo('Saídas (produtos)', chips(a.Saidas, 'fa-arrow-right-from-bracket')) +
      campo('Sistemas', chips(a.Sistemas, 'fa-desktop')) +
      campo('Base normativa', a.Base_Normativa ? chips(a.Base_Normativa, 'fa-scale-balanced') : null) + '</dl></div>' +
      secVinculos('Atividade', cod) +
      '</div><aside>' +
      (sp2 ? '<div class="pp-card"><h3><i class="fas fa-arrow-turn-up" aria-hidden="true"></i> Subprocesso pai</h3>' +
        '<a class="proc-card" href="#/sp/' + encodeURIComponent(sp2.Codigo) + '"><div class="topo"><div><span class="cod">' + esc(sp2.Codigo) + '</span><div class="nome" style="font-size:var(--fs-sm)">' + esc(sp2.Nome) + '</div></div></div></a></div>' : '') +
      '</aside></div>';
  }
  function naoEncontrado(tipo, cod) {
    return breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: tipo + ' não encontrado' }]) +
      '<div class="pp-card"><h3>' + esc(tipo) + ' não encontrado</h3><p style="font-size:var(--fs-sm)">O código <strong>' +
      esc(cod) + '</strong> não existe na base atual. Verifique a planilha ou volte ao <a href="#/catalogo">catálogo</a>.</p></div>';
  }

  /* ── TELAS: documentos · riscos · indicadores · diário ────────────── */
  var filtroDoc = { tipo: '', q: '' };
  function renderDocumentos() {
    var el = $('#viewDocumentos');
    var tipos = {};
    DADOS.docs.forEach(function (x) { if (x.Tipo_Documento) tipos[x.Tipo_Documento] = 1; });
    var lista = DADOS.docs.filter(function (x) {
      if (filtroDoc.tipo && x.Tipo_Documento !== filtroDoc.tipo) return false;
      if (filtroDoc.q && (x.ID + ' ' + x.Titulo).toLowerCase().indexOf(filtroDoc.q.toLowerCase()) < 0) return false;
      return true;
    });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Repositório de documentos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">POPs, manuais, atas, diagramas BPMN e relatórios produzidos no mapeamento, vinculados de forma rastreável a macroprocessos, processos, subprocessos e atividades.</p>' +
      '<div class="pp-filtros"><label class="sr-only" for="fTipoDoc">Tipo de documento</label>' +
      '<select id="fTipoDoc"><option value="">Todos os tipos</option>' +
      Object.keys(tipos).sort().map(function (t) { return '<option' + (filtroDoc.tipo === t ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join('') + '</select>' +
      '<label class="sr-only" for="fBuscaDoc">Buscar documento</label>' +
      '<input type="search" id="fBuscaDoc" placeholder="Buscar por título…" value="' + esc(filtroDoc.q) + '">' +
      '<span class="pp-muted" style="font-size:var(--fs-sm)">' + lista.length + ' de ' + DADOS.docs.length + '</span></div>' +
      '<div class="pp-card"><div class="pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th><th>Documento</th><th>Vinculado a</th><th>Data</th><th>Situação</th></tr></thead><tbody>' +
      (lista.length ? lista.map(function (x) {
        var tit = x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '<span class="sr-only"> (abre em nova aba)</span></a>' : esc(x.Titulo);
        return '<tr><td class="cod">' + esc(x.ID) + '</td><td><strong>' + tit + '</strong><div class="pp-muted" style="font-size:var(--fs-xs)">' +
          esc(x.Tipo_Documento || '') + (x.Versao ? ' · v' + esc(x.Versao) : '') + '</div></td>' +
          '<td>' + linkVinculo(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td><td>' + fmtData(x.Data) + '</td><td>' + esc(x.Situacao || '—') + '</td></tr>';
      }).join('') : '<tr><td colspan="5" class="pp-vazio">Nenhum documento corresponde aos filtros.</td></tr>') +
      '</tbody></table></div></div>';
    $('#fTipoDoc').onchange = function () { filtroDoc.tipo = this.value; renderDocumentos(); };
    $('#fBuscaDoc').oninput = function () { filtroDoc.q = this.value; renderDocumentos(); $('#fBuscaDoc').focus(); };
  }
  function renderRiscos() {
    var el = $('#viewRiscos');
    var celulas = '';
    for (var imp = 5; imp >= 1; imp--) {
      celulas += '<div class="cab">' + imp + '</div>';
      for (var prob = 1; prob <= 5; prob++) {
        var nivel = prob * imp;
        var cls = 'n-' + slug(classeRisco(nivel));
        var pins = DADOS.riscos.filter(function (r) { return r.Probabilidade_1a5 === prob && r.Impacto_1a5 === imp; })
          .map(function (r) {
            return '<button type="button" class="risco-pin" title="' + esc(r.Descricao_Risco) + '" data-alvo="risco-' + esc(r.ID) + '">' + esc(r.ID.replace('R-', '')) + '</button>';
          }).join('');
        celulas += '<div class="cel ' + cls + '">' + pins + '</div>';
      }
    }
    celulas += '<div class="cab"></div>';
    for (var pr = 1; pr <= 5; pr++) celulas += '<div class="cab">' + pr + '</div>';
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Radar de riscos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Riscos identificados durante o mapeamento, vinculados ao nível em que foram observados. Nível = Probabilidade × Impacto (matriz 5×5).</p>' +
      '<div class="pp-card"><h3><i class="fas fa-border-all" aria-hidden="true"></i> Matriz 5×5 (Impacto ↑ × Probabilidade →)</h3>' +
      '<div class="matriz" role="img" aria-label="Matriz de riscos cinco por cinco">' + celulas + '</div>' +
      '<div class="matriz-legenda">' + ['Baixo', 'Moderado', 'Alto', 'Extremo'].map(function (c) { return tagNivel(c); }).join('') +
      '<span class="pp-muted">Clique em um risco para ver os detalhes na tabela.</span></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Todos os riscos</h3>' + tabelaRiscosHtml(DADOS.riscos, true) + '</div>';
    $all('#viewRiscos .risco-pin').forEach(function (b) {
      b.addEventListener('click', function () {
        var alvo = d.getElementById(b.getAttribute('data-alvo'));
        if (alvo) { alvo.scrollIntoView({ behavior: 'smooth', block: 'center' }); alvo.style.background = 'var(--blue-warm-vivid-5)'; setTimeout(function () { alvo.style.background = ''; }, 1600); }
      });
    });
  }
  function renderIndicadores() {
    var el = $('#viewIndicadores');
    var atingidas = DADOS.inds.filter(function (x) { return x._sit === 'Meta atingida'; }).length;
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Indicadores de desempenho</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Medição dos processos (CBOK 4.0 — fase 5, medir o sucesso): ' +
      atingidas + ' de ' + DADOS.inds.length + ' indicadores com meta atingida.</p>' +
      '<div class="pp-card">' + tabelaIndsHtml(DADOS.inds, true) + '</div>';
  }
  var filtroDiario = '';
  function renderDiario() {
    var el = $('#viewDiario');
    var lista = filtroDiario ? DADOS.diario.filter(function (e) { return e.Processo === filtroDiario; }) : DADOS.diario;
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Diário de mapeamento</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Registro rastreável do trabalho de gestão de processos: reuniões, oficinas, entrevistas, validações, decisões e entregas — cada um com insumos, entregáveis e evidências (CBOK 4.0 · PMBOK).</p>' +
      '<div class="pp-filtros"><label class="sr-only" for="fProcDiario">Filtrar por processo</label>' +
      '<select id="fProcDiario"><option value="">Todos os processos</option>' +
      DADOS.procs.filter(function (p) { return (IDX.diarioPorProc[p.Codigo] || []).length; }).map(function (p) {
        return '<option value="' + esc(p.Codigo) + '"' + (filtroDiario === p.Codigo ? ' selected' : '') + '>' + esc(p.Codigo + ' — ' + p.Nome) + '</option>';
      }).join('') + '</select>' +
      '<span class="pp-muted" style="font-size:var(--fs-sm)">' + lista.length + ' registros</span></div>' +
      '<div class="pp-card">' + timelineHtml(lista) + '</div>';
    $('#fProcDiario').onchange = function () { filtroDiario = this.value; renderDiario(); };
  }

  /* ── TELA: metodologia ────────────────────────────────────────────── */
  function renderMetodologia() {
    $('#viewMetodologia').innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Metodologia</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-rotate" aria-hidden="true"></i> Ciclo de vida BPM (BPM CBOK 4.0)</h3>' +
      '<p style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Cada processo da carteira percorre as cinco fases do ciclo de vida BPM. A fase atual de cada processo aparece na respectiva ficha.</p>' +
      '<ol class="ciclo">' +
      '<li><h4>Alinhamento à estratégia e metas</h4><p>Priorização da carteira e vínculo do processo aos objetivos institucionais.</p></li>' +
      '<li><h4>Arquitetar mudanças</h4><p>Modelagem (AS-IS), análise, desenho do estado futuro (TO-BE) e definição da medição.</p></li>' +
      '<li><h4>Desenvolver iniciativas</h4><p>Planos de implantação, capacitação, mudanças e tecnologia (visão PMBOK do projeto).</p></li>' +
      '<li><h4>Implementar mudanças</h4><p>Execução dos planos, publicação de procedimentos e estabilização.</p></li>' +
      '<li><h4>Medir o sucesso</h4><p>Monitoramento por indicadores e melhoria contínua (novo giro do ciclo).</p></li>' +
      '</ol></div>' +
      '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M9)</h3>' +
      '<p style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Roteiro-padrão de cada projeto de mapeamento, do primeiro contato com a área até a publicação no repositório:</p>' +
      '<ul class="marcos">' + MARCOS_ROTULOS.map(function (r) { return '<li class="feito"><span>' + esc(r) + '</span><i class="fas fa-check-circle" aria-hidden="true"></i></li>'; }).join('') + '</ul></div>' +
      '<div class="pp-card"><h3><i class="fas fa-layer-group" aria-hidden="true"></i> Hierarquia e tipos de processos</h3>' +
      '<dl class="ficha-dl">' +
      campo('Hierarquia (níveis de modelo)', 'Macroprocesso → Processo → Subprocesso → Atividade. Cada nível tem ficha própria, diagrama BPMN e vínculos de documentos, riscos e indicadores.', true) +
      campo('Processos finalísticos (primários)', 'Entregam valor diretamente ao cliente/beneficiário — a razão de ser da organização.') +
      campo('Processos de suporte', 'Sustentam os finalísticos (contratações, pessoas, TI), sem entregar valor direto ao beneficiário.') +
      campo('Processos gerenciais', 'Medem, monitoram e controlam a atuação institucional (estratégia, riscos, governança).') +
      campo('Dono do processo (process owner)', 'Responsável fim a fim pelo desempenho do processo — papel central da disciplina de BPM.') +
      '</dl></div>' +
      '<div class="pp-card"><h3><i class="fas fa-database" aria-hidden="true"></i> Como este painel é alimentado</h3>' +
      '<dl class="ficha-dl">' +
      campo('1. Google Sheets (recomendado)', 'Importe a planilha-modelo para o Google Sheets, compartilhe como “qualquer pessoa com o link pode ver” e informe o ID em <code>PAINEL_CONFIG.googleSheetId</code> (index.html). O painel passa a ler em tempo real — mesmo padrão do Painel do PTD.', true) +
      campo('2. Planilha no repositório', 'Sem Google Sheets, o painel lê <code>data/painel-processos-dados.xlsx</code> publicado junto com o site (GitHub Pages).', true) +
      campo('3. Dados embutidos', 'Como reserva (inclusive offline/file://), usa <code>js/dados.js</code>, gerado por <code>scripts/planilha_para_js.py</code>.', true) +
      '</dl></div>' +
      '<div class="pp-card"><h3><i class="fas fa-book" aria-hidden="true"></i> Referências</h3>' +
      '<ul style="font-size:var(--fs-sm);padding-left:1.2rem">' +
      '<li>ABPMP. <strong>BPM CBOK — Guia para o Gerenciamento de Processos de Negócio: Corpo Comum de Conhecimento</strong>, versão 4.0 (hierarquia e tipos de processos, ciclo de vida BPM, papéis, SIPOC, medição de desempenho e repositório de processos).</li>' +
      '<li>PMI. <strong>Guia PMBOK</strong> — gestão dos projetos de mapeamento: termo de abertura, escopo, partes interessadas, riscos, entregáveis e lições aprendidas.</li>' +
      '<li>Gov.br. <strong>Design System (govbr-ds v4)</strong> — padrão visual e de acessibilidade deste painel.</li></ul>' +
      '<p class="pp-aviso" style="margin-top:var(--sp2)"><strong>Aviso:</strong> os dados atualmente exibidos são <strong>fictícios</strong>, criados para demonstração da estrutura do painel.</p></div>';
  }

  /* ── TELA: busca global ───────────────────────────────────────────── */
  function renderBusca(q) {
    var el = $('#viewBusca');
    var ql = q.toLowerCase();
    function bate(txt) { return String(txt || '').toLowerCase().indexOf(ql) >= 0; }
    function grupo(titulo, itens, fmt) {
      if (!itens.length) return '';
      return '<div class="pp-card"><h3>' + titulo + ' (' + itens.length + ')</h3>' + itens.map(fmt).join('') + '</div>';
    }
    var r = {
      mp: DADOS.macros.filter(function (m) { return bate(m.Codigo) || bate(m.Nome) || bate(m.Descricao); }),
      p: DADOS.procs.filter(function (p) { return bate(p.Codigo) || bate(p.Nome) || bate(p.Descricao); }),
      sp: DADOS.subs.filter(function (s) { return bate(s.Codigo) || bate(s.Nome) || bate(s.Descricao); }),
      a: DADOS.ativs.filter(function (a) { return bate(a.Codigo) || bate(a.Nome) || bate(a.Descricao); }),
      doc: DADOS.docs.filter(function (x) { return bate(x.ID) || bate(x.Titulo); }),
      reg: DADOS.diario.filter(function (e) { return bate(e.Titulo) || bate(e.Descricao); })
    };
    var total = r.mp.length + r.p.length + r.sp.length + r.a.length + r.doc.length + r.reg.length;
    function linha(href, cod, nome, extra) {
      return '<div class="doc-item"><i class="fas fa-arrow-right fa-stack-ico" aria-hidden="true"></i><div>' +
        '<div class="tit"><a href="' + href + '"><span class="cod">' + esc(cod) + '</span> ' + esc(nome) + '</a></div>' +
        (extra ? '<div class="meta">' + extra + '</div>' : '') + '</div></div>';
    }
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Busca' }]) +
      '<div class="pp-sec-h" style="margin-top:0"><h2>Resultados para “' + esc(q) + '”</h2><div class="linha" aria-hidden="true"></div></div>' +
      (total ? '' : '<p class="pp-vazio">Nada encontrado. Tente outro termo ou navegue pelo <a href="#/catalogo">catálogo</a>.</p>') +
      grupo('Macroprocessos', r.mp, function (m) { return linha('#/mp/' + encodeURIComponent(m.Codigo), m.Codigo, m.Nome, esc(m.Categoria)); }) +
      grupo('Processos', r.p, function (p) { return linha('#/p/' + encodeURIComponent(p.Codigo), p.Codigo, p.Nome, esc(p.Status_Mapeamento) + ' · ' + p.Percentual + '%'); }) +
      grupo('Subprocessos', r.sp, function (s) { return linha('#/sp/' + encodeURIComponent(s.Codigo), s.Codigo, s.Nome, ''); }) +
      grupo('Atividades', r.a, function (a) { return linha('#/a/' + encodeURIComponent(a.Codigo), a.Codigo, a.Nome, esc(a.Responsavel_Ator || '')); }) +
      grupo('Documentos', r.doc, function (x) {
        return '<div class="doc-item"><i class="fas fa-file fa-stack-ico" aria-hidden="true"></i><div><div class="tit">' +
          (x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '</a>' : esc(x.Titulo)) +
          '</div><div class="meta">' + linkVinculo(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</div></div></div>';
      }) +
      grupo('Diário de mapeamento', r.reg, function (e) {
        return linha('#/p/' + encodeURIComponent(e.Processo), e.Processo, e.Titulo || '(registro)', fmtData(e.Data) + ' · ' + esc(e.Tipo || ''));
      });
  }

  /* ── inicialização ────────────────────────────────────────────────── */
  function posCarga() {
    var chip = $('#syncChip'); if (chip) chip.textContent = FONTE;
    var c;
    if ((c = $('#cntCatalogo'))) c.textContent = DADOS.procs.length;
    if ((c = $('#cntDocumentos'))) c.textContent = DADOS.docs.length;
    if ((c = $('#cntRiscos'))) c.textContent = DADOS.riscos.length;
    if ((c = $('#cntIndicadores'))) c.textContent = DADOS.inds.length;
    if ((c = $('#cntDiario'))) c.textContent = DADOS.diario.length;
    montarAlertas();
    ligarAcoesCabecalho();
    if (window.PPUI) PPUI.setMenuSections([
      { rotulo: 'Início · Cadeia de Valor', href: '#/', icone: 'fa-house' },
      { rotulo: 'Catálogo de processos', href: '#/catalogo', icone: 'fa-layer-group' },
      { rotulo: 'Documentos', href: '#/documentos', icone: 'fa-folder-open' },
      { rotulo: 'Radar de riscos', href: '#/riscos', icone: 'fa-shield-halved' },
      { rotulo: 'Indicadores', href: '#/indicadores', icone: 'fa-chart-line' },
      { rotulo: 'Diário de mapeamento', href: '#/diario', icone: 'fa-timeline' },
      { rotulo: 'Metodologia (CBOK · PMBOK)', href: '#/metodologia', icone: 'fa-book' }
    ]);
    rota();
  }
  function iniciar() {
    var v = $('#viewInicio');
    if (v) v.innerHTML = '<div class="pp-loading"><i class="fas fa-circle-notch" aria-hidden="true"></i> Carregando dados do painel…</div>';
    carregarDados().then(posCarga).catch(function (e) {
      console.error(e);
      if (v) v.innerHTML = '<div class="pp-aviso"><strong>Não foi possível carregar os dados.</strong> ' +
        'Verifique se data/painel-processos-dados.xlsx está publicado (ou gere js/dados.js com scripts/planilha_para_js.py). Detalhe: ' + esc(e.message) + '</div>';
    });
  }
  window.addEventListener('hashchange', rota);
  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
