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
    abas: ['Macroprocessos', 'Processos', 'Subprocessos', 'Atividades', 'Tarefas',
           'Documentos', 'Riscos', 'Indicadores', 'Diario_Mapeamento',
           'Jornada', 'Repositorio', 'NUGEP', 'Glossario', 'FAQ', 'Parametros']
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
      tarefas: pega('Tarefas'),
      docs: pega('Documentos'),
      riscos: pega('Riscos'),
      inds: pega('Indicadores'),
      diario: pega('Diario_Mapeamento'),
      jornada: pega('Jornada'),
      repo: pega('Repositorio'),
      nugep: pega('NUGEP'),
      glossario: pega('Glossario'),
      faq: pega('FAQ'),
      parametros: pega('Parametros')
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
    dd.jornada.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.repo.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.nugep.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.faq.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.glossario.sort(function (a, b) { return String(a.Termo || '').localeCompare(String(b.Termo || ''), 'pt-BR'); });
    dd.params = {};
    dd.parametros.forEach(function (x) { if (x.Chave) dd.params[x.Chave] = x.Valor || ''; });

    var idx = { mp: {}, p: {}, sp: {}, a: {}, t: {}, procsPorMacro: {}, subsPorProc: {}, ativsPorSub: {}, tarefasPorAtiv: {},
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
    dd.tarefas.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.tarefas.forEach(function (t) {
      idx.t[t.Codigo] = t;
      (idx.tarefasPorAtiv[t.Atividade] = idx.tarefasPorAtiv[t.Atividade] || []).push(t);
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
  var NIVEL_PREFIXO = { 'Macroprocesso': 'mp', 'Processo': 'p', 'Subprocesso': 'sp',
    'Atividade': 'a', 'Tarefa': 't' };
  var NIVEL_ROTULO = { 'Macroprocesso': 'Macroprocesso', 'Processo': 'Processo de negócio',
    'Subprocesso': 'Subprocesso', 'Atividade': 'Atividade', 'Tarefa': 'Tarefa' };
  function nivelRotulo(n) { return NIVEL_ROTULO[n] || n; }
  function rotaDe(nivel, codigo) {
    var pre = NIVEL_PREFIXO[nivel];
    return pre ? '#/' + pre + '/' + encodeURIComponent(codigo) : '#/';
  }
  function nomeDe(nivel, codigo) {
    var it = nivel === 'Macroprocesso' ? IDX.mp[codigo] : nivel === 'Processo' ? IDX.p[codigo]
      : nivel === 'Subprocesso' ? IDX.sp[codigo] : nivel === 'Tarefa' ? IDX.t[codigo] : IDX.a[codigo];
    return it ? it.Nome : codigo;
  }
  function linkVinculo(nivel, codigo) {
    return '<a href="' + rotaDe(nivel, codigo) + '"><span class="cod">' + esc(codigo) +
      '</span> ' + esc(nomeDe(nivel, codigo)) + '</a> <span class="pp-muted">(' + esc(nivelRotulo(nivel)) + ')</span>';
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
  function urlDrive(u) {                     // link de compartilhamento → imagem exibível
    var m = String(u || '').match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]{20,})/);
    return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w2000' : null;
  }
  function diagramaHtml(caminho, titulo) {
    if (!caminho) return '<p class="pp-vazio">Diagrama BPMN ainda não publicado para este item. Informe a URL da imagem exportada do Bizagi (ou um caminho do repositório) na coluna Imagem_Bizagi da planilha.</p>';
    var href = esc(caminho);
    var src = esc(urlDrive(caminho) || caminho);
    return '<figure class="diagrama-frame"><a href="' + href + '" target="_blank" rel="noopener" title="Abrir o diagrama no link original (nova aba)">' +
      '<img src="' + src + '" alt="Diagrama BPMN (Bizagi) — ' + esc(titulo) + '" loading="lazy" ' +
      'onerror="this.closest(&quot;figure&quot;).classList.add(&quot;sem-imagem&quot;)"></a>' +
      '<figcaption class="diagrama-fallback"><i class="fas fa-diagram-project" aria-hidden="true"></i> A pré-visualização não pôde ser carregada aqui — abra o diagrama pelo botão abaixo.</figcaption></figure>' +
      '<div class="diagrama-acoes"><a class="br-button secondary small" href="' + href +
      '" target="_blank" rel="noopener"><i class="fas fa-up-right-from-square" aria-hidden="true"></i>&nbsp;Abrir diagrama no link publicado<span class="sr-only"> (abre em nova aba)</span></a></div>';
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
    return '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th>' +
      (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Risco</th><th>P</th><th>I</th><th>P×I</th><th>Nível</th><th>Resposta</th><th>Status</th></tr></thead><tbody>' +
      riscos.map(function (r) {
        return '<tr id="risco-' + esc(r.ID) + '"><td class="cod">' + esc(r.ID) + '</td>' +
          (comVinculo ? '<td>' + linkVinculo(r.Vinculo_Nivel, r.Vinculo_Codigo) + '</td>' : '') +
          '<td>' + esc(r.Descricao_Risco) +
          (r.Controles_Tratamento ? '<div class="pp-muted" style="font-size:var(--fs-sm)">Tratamento: ' + esc(r.Controles_Tratamento) + '</div>' : '') +
          '</td><td>' + r.Probabilidade_1a5 + '</td><td>' + r.Impacto_1a5 + '</td><td><strong>' + r._nivel +
          '</strong></td><td>' + tagNivel(r._classe) + '</td><td>' + esc(r.Resposta || '—') +
          '</td><td>' + esc(r.Status || '—') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  function tabelaIndsHtml(inds, comVinculo) {
    if (!inds.length) return '<p class="pp-vazio">Nenhum indicador vinculado.</p>';
    return '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th>' +
      (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Indicador</th><th>Meta</th><th>Resultado</th><th>Situação</th><th>Periodicidade</th><th>Última medição</th></tr></thead><tbody>' +
      inds.map(function (x) {
        var cls = x._sit === 'Meta atingida' ? 'sit-ok' : (x._sit === 'Sem medição' || x._sit === 'Sem meta') ? 'sit-neutra' : 'sit-ruim';
        var un = x.Unidade ? ' ' + esc(x.Unidade) : '';
        return '<tr><td class="cod">' + esc(x.ID) + '</td>' +
          (comVinculo ? '<td>' + linkVinculo(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' : '') +
          '<td><strong>' + esc(x.Nome) + '</strong>' +
          (x.Descricao_Formula ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(x.Descricao_Formula) + '</div>' : '') +
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
        (e.Participantes ? '<p class="pp-muted" style="font-size:var(--fs-sm)"><i class="fas fa-users" aria-hidden="true"></i> ' + esc(e.Participantes) + '</p>' : '') +
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
  var ROTAS_ABA = { inicio: '#/', catalogo: '#/catalogo', dashboard: '#/dashboard', repositorio: '#/repositorio',
    documentos: '#/documentos', riscos: '#/riscos', indicadores: '#/indicadores',
    diario: '#/diario', nugep: '#/nugep', glossario: '#/glossario', faq: '#/faq' };
  var MAIS_ITENS = { diario: 'Diário', repositorio: 'Repositório', nugep: 'NUGEP',
    glossario: 'Glossário', faq: 'FAQ' };
  function mostrarPainel(id) {
    $all('#mainTabContent > .tab-panel').forEach(function (p) {
      var ativo = p.id === 'panel-' + id;
      p.classList.toggle('active', ativo);
      p.hidden = !ativo;
    });
    $all('.tab-nav > ul > .tab-item > [data-rota]').forEach(function (b) {
      var ativo = b.getAttribute('data-painel') === id ||
        (id === 'detalhe' && b.getAttribute('data-painel') === 'catalogo') ||
        (id === 'busca' && b.getAttribute('data-painel') === 'inicio');
      b.setAttribute('aria-selected', ativo ? 'true' : 'false');
      var li = b.closest('.tab-item'); if (li) li.classList.toggle('active', ativo);
    });
    $all('#maisMenu [data-rota]').forEach(function (b) {
      var ativo = b.getAttribute('data-painel') === id;
      b.classList.toggle('ativo', ativo);
      if (ativo) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
    });
    var maisBtn = $('#tab-mais');
    if (maisBtn) {
      var sub = MAIS_ITENS[id], temAtivo = !!sub;
      maisBtn.setAttribute('aria-selected', temAtivo ? 'true' : 'false');
      var maisLi = maisBtn.closest('.tab-item'); if (maisLi) maisLi.classList.toggle('active', temAtivo);
      var lf = $('#maisLabelFull'), ls = $('#maisLabelShort');
      if (lf) lf.textContent = temAtivo ? sub : 'Mais';
      if (ls) ls.textContent = temAtivo ? sub : 'Mais';
      maisBtn.setAttribute('aria-label', temAtivo ? ('Mais seções do painel — atual: ' + sub) : 'Mais seções do painel');
    }
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
    else if (h === '#/dashboard') { renderDashboard(); mostrarPainel('dashboard'); }
    else if (h === '#/repositorio' || h === '#/metodologia') { renderRepositorio(); mostrarPainel('repositorio'); }
    else if (h === '#/nugep') { renderNugep(); mostrarPainel('nugep'); }
    else if (h === '#/glossario') { renderGlossario(); mostrarPainel('glossario'); }
    else if (h === '#/faq') { renderFaq(); mostrarPainel('faq'); }
    else if ((m = h.match(/^#\/busca\?q=(.*)$/))) { renderBusca(decodeURIComponent(m[1])); mostrarPainel('busca'); }
    else if ((m = h.match(/^#\/(mp|p|sp|a|t)\/(.+)$/))) { renderDetalhe(m[1], decodeURIComponent(m[2])); mostrarPainel('detalhe'); }
    else { renderInicio(); mostrarPainel('inicio'); }
    if (jaNavegou) {
      try { window.scrollTo({ top: 0, behavior: 'auto' }); } catch (e) { window.scrollTo(0, 0); }
      var h1 = d.querySelector('#mainTabContent .tab-panel.active h1, #mainTabContent .tab-panel.active h2');
      if (h1) {
        if (!h1.hasAttribute('tabindex')) h1.setAttribute('tabindex', '-1');
        h1.focus({ preventScroll: true });
      }
    }
    jaNavegou = true;
  }
  var jaNavegou = false;
  d.addEventListener('click', function (ev) {
    var b = ev.target.closest('.tab-nav [data-rota]');
    if (!b) return;
    location.hash = b.getAttribute('data-rota');
    var dd = b.closest('.dd-target');
    if (dd) {
      dd.hidden = true;
      var trig = $('[data-target="' + dd.id + '"]');
      if (trig) trig.setAttribute('aria-expanded', 'false');
    }
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
  var MP_ICONES = { 'MP-01': 'fa-bullseye', 'MP-02': 'fa-shield-halved', 'MP-03': 'fa-seedling',
    'MP-04': 'fa-droplet', 'MP-05': 'fa-water', 'MP-06': 'fa-file-contract',
    'MP-07': 'fa-users', 'MP-08': 'fa-laptop-code' };
  function statsMacro(cod) {
    var ps = IDX.procsPorMacro[cod] || [];
    if (!ps.length) return 'sem processos cadastrados';
    var media = Math.round(ps.reduce(function (s, p) { return s + p.Percentual; }, 0) / ps.length);
    return ps.length + (ps.length === 1 ? ' processo' : ' processos') + ' · ' + media + '% mapeado';
  }
  function blocoCadeia(titulo, classe, icone, itens) {
    return '<div class="cv-bloco ' + classe + '"><div class="cv-titulo"><i class="fas ' + icone +
      '" aria-hidden="true"></i><span>' + titulo + '</span><span class="cv-qtd">' + itens.length + '</span></div><ul>' +
      itens.map(function (m) {
        return '<li><a href="#/mp/' + encodeURIComponent(m.Codigo) + '">' +
          '<span class="cv-ico"><i class="fas ' + (MP_ICONES[m.Codigo] || 'fa-diagram-project') + '" aria-hidden="true"></i></span>' +
          '<span class="cv-tx"><span class="cod">' + esc(m.Codigo) + '</span><span class="nome">' + esc(m.Nome) +
          '</span><span class="cv-meta">' + statsMacro(m.Codigo) + '</span></span></a></li>';
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
      '<section class="pp-hero">' +
      '<h1>Cadeia de valor e mapeamento de processos da Codevasf</h1>' +
      '<p>Consulte a hierarquia completa — do macroprocesso à atividade — com fichas, diagramas BPMN (Bizagi), ' +
      'documentos, riscos, indicadores e o registro rastreável de cada mapeamento realizado.</p>' +
      '</section>' +
      '<div class="kpi-grid">' +
      '<div class="kpi"><span class="num">' + DADOS.macros.length + '</span><span class="lbl">Macroprocessos</span><span class="sub">' + procs.length + ' processos · ' + DADOS.subs.length + ' subprocessos · ' + DADOS.ativs.length + ' atividades · ' + DADOS.tarefas.length + ' tarefas</span></div>' +
      '<div class="kpi ok"><span class="num">' + concl + '</span><span class="lbl">Mapeamentos concluídos</span><span class="sub">' + andamento + ' em andamento</span></div>' +
      '<div class="kpi"><span class="num">' + media + '%</span><span class="lbl">Avanço médio</span><span class="sub">do mapeamento da carteira</span></div>' +
      '<div class="kpi ' + (criticos ? 'erro' : 'ok') + '"><span class="num">' + criticos + '</span><span class="lbl">Riscos críticos abertos</span><span class="sub">nível Alto ou Extremo</span></div>' +
      '<div class="kpi"><span class="num">' + docsVig + '</span><span class="lbl">Documentos vigentes</span><span class="sub">' + DADOS.diario.length + ' registros no diário</span></div>' +
      '</div>' +
      '<section class="pp-sec" id="sec-cadeia"><div class="pp-sec-h"><h2>Cadeia de Valor Integrada</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="cadeia">' +
      '<aside class="cv-aside cv-missao"><h3><i class="fas fa-flag" aria-hidden="true"></i> Missão</h3><p>' + esc(INSTITUCIONAL.missao) + '</p><h3><i class="fas fa-eye" aria-hidden="true"></i> Visão</h3><p>' + esc(INSTITUCIONAL.visao) + '</p></aside>' +
      '<div class="cv-centro">' + blocoCadeia('Macroprocessos Gerenciais', 'cat-gerencial', 'fa-compass', ger) +
      blocoCadeia('Macroprocessos Finalísticos — entrega de valor à sociedade', 'cat-finalistico', 'fa-hand-holding-heart', fin) + '</div>' +
      '<aside class="cv-aside cv-proposito"><h3><i class="fas fa-hand-holding-heart" aria-hidden="true"></i> Propósito</h3><p>' + esc(INSTITUCIONAL.proposito) + '</p></aside>' +
      '<div class="cv-suporte">' + blocoCadeia('Macroprocessos de Suporte', 'cat-suporte', 'fa-gears', sup) + '</div>' +
      '<div class="cv-valores"><strong><i class="fas fa-gem" aria-hidden="true"></i> Valores</strong>' + INSTITUCIONAL.valores.map(function (v) { return '<span class="cv-chip">' + esc(v) + '</span>'; }).join('') + '</div>' +
      '</div><p class="pp-muted" style="margin-top:var(--sp2);font-size:var(--fs-sm)">Clique em um macroprocesso para abrir a ficha e navegar até os processos de negócio, subprocessos, atividades e tarefas.</p></section>';
  }

  /* ── TELA: catálogo ───────────────────────────────────────────────── */
  var filtroCat = { macro: '', status: '', q: '' };
  function cardProcesso(p) {
    return '<a class="proc-card" href="#/p/' + encodeURIComponent(p.Codigo) + '">' +
      '<div class="topo"><div><span class="cod" style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(p.Codigo) + '</span>' +
      '<div class="nome">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
      '<div class="pp-muted" style="font-size:var(--fs-sm);margin-top:4px">' + esc(p.Area_Responsavel || '') +
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
        campo('Beneficiários', chips(m.Clientes_Beneficiarios)) +
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
            '<div class="topo"><div><span style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(p.Codigo) + '</span>' +
            '<div class="nome" style="font-size:var(--fs-sm)">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
            '<div class="rodape">' + barraPct(p.Percentual) + '</div></a>';
        }).join('') : '<p class="pp-vazio">Nenhum processo cadastrado.</p>') + '</div></aside></div>';
      return;
    }
    if (tipo === 'p') {
      var p = IDX.p[cod];
      if (!p) { el.innerHTML = naoEncontrado('Processo de negócio', cod); return; }
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
        '<div class="col"><h4>Beneficiários</h4><ul>' + (listar(p.Beneficiarios || p.Clientes).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '</div></div>' +
        '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M9)</h3>' + marcosHtml(p) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(p.Imagem_Bizagi, p.Nome) + '</div>' +
        secVinculos('Processo', cod) +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos</h3>' +
        (subs.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Subprocesso</th><th>Entregas</th><th></th></tr></thead><tbody>' +
          subs.map(function (s) {
            return '<tr><td class="cod">' + esc(s.Codigo) + '</td><td><a href="#/sp/' + encodeURIComponent(s.Codigo) + '"><strong>' + esc(s.Nome) + '</strong></a>' +
              (s.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(s.Descricao) + '</div>' : '') + '</td>' +
              '<td style="font-size:var(--fs-sm)">' + (listar(s.Entregas).map(esc).join('; ') || '—') + '</td>' +
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
        (ativs.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>#</th><th>Atividade</th><th>Responsável (ator)</th><th>Entradas</th><th>Saídas</th><th>Prazo</th></tr></thead><tbody>' +
          ativs.map(function (a, i) {
            return '<tr data-link><td>' + (i + 1) + '</td><td><a href="#/a/' + encodeURIComponent(a.Codigo) + '"><strong>' + esc(a.Nome) + '</strong></a>' +
              '<div class="cod">' + esc(a.Codigo) + '</div></td><td style="font-size:var(--fs-sm)">' + esc(a.Responsavel_Ator || '—') + '</td>' +
              '<td style="font-size:var(--fs-sm)">' + (listar(a.Entradas).map(esc).join('; ') || '—') + '</td>' +
              '<td style="font-size:var(--fs-sm)">' + (listar(a.Saidas).map(esc).join('; ') || '—') + '</td>' +
              '<td style="font-size:var(--fs-sm);white-space:nowrap">' + esc(a.Prazo_Padrao || '—') + '</td></tr>';
          }).join('') + '</tbody></table></div>' : '<p class="pp-vazio">Nenhuma atividade cadastrada.</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(s.Imagem_Bizagi, s.Nome) + '</div>' +
        secVinculos('Subprocesso', cod) +
        '</div><aside>' +
        (pp ? '<div class="pp-card"><h3><i class="fas fa-arrow-turn-up" aria-hidden="true"></i> Processo de negócio (pai)</h3>' +
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
    if (tipo === 'a') {
    var a = IDX.a[cod];
    if (!a) { el.innerHTML = naoEncontrado('Atividade', cod); return; }
    var sp2 = IDX.sp[a.Subprocesso]; var p2 = sp2 && IDX.p[sp2.Processo]; var mp2 = p2 && IDX.mp[p2.Macroprocesso];
    var tf3 = IDX.tarefasPorAtiv[cod] || [];
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
      '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(a.Imagem_Bizagi, a.Nome) + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Tarefas (menor unidade de trabalho — CBOK 4.0)</h3>' +
      (tf3.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Tarefa</th><th>Tipo</th><th>Duração</th></tr></thead><tbody>' +
        tf3.map(function (t) {
          return '<tr data-link><td class="cod">' + esc(t.Codigo) + '</td><td><a href="#/t/' + encodeURIComponent(t.Codigo) + '"><strong>' + esc(t.Nome) + '</strong></a>' +
            (t.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(t.Descricao) + '</div>' : '') + '</td>' +
            '<td style="font-size:var(--fs-sm)">' + esc(t.Tipo_Tarefa || '—') + '</td>' +
            '<td style="font-size:var(--fs-sm);white-space:nowrap">' + esc(t.Duracao_Estimada || '—') + '</td></tr>';
        }).join('') + '</tbody></table></div>' : '<p class="pp-vazio">Nenhuma tarefa cadastrada para esta atividade.</p>') + '</div>' +
      secVinculos('Atividade', cod) +
      '</div><aside>' +
      (sp2 ? '<div class="pp-card"><h3><i class="fas fa-arrow-turn-up" aria-hidden="true"></i> Subprocesso (pai)</h3>' +
        '<a class="proc-card" href="#/sp/' + encodeURIComponent(sp2.Codigo) + '"><div class="topo"><div><span class="cod">' + esc(sp2.Codigo) + '</span><div class="nome" style="font-size:var(--fs-sm)">' + esc(sp2.Nome) + '</div></div></div></a></div>' : '') +
      '</aside></div>';
    $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        if (ev.target.closest('a')) return;
        var lk = tr.querySelector('a'); if (lk) location.hash = lk.getAttribute('href');
      });
    });
    return;
    }
    // tipo === 't' — ficha da tarefa
    var t = IDX.t[cod];
    if (!t) { el.innerHTML = naoEncontrado('Tarefa', cod); return; }
    var a3 = IDX.a[t.Atividade]; var sp3 = a3 && IDX.sp[a3.Subprocesso];
    var p3 = sp3 && IDX.p[sp3.Processo]; var mp3 = p3 && IDX.mp[p3.Macroprocesso];
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp3 ? [{ rotulo: mp3.Codigo, href: '#/mp/' + encodeURIComponent(mp3.Codigo) }] : [])
        .concat(p3 ? [{ rotulo: p3.Codigo, href: '#/p/' + encodeURIComponent(p3.Codigo) }] : [])
        .concat(sp3 ? [{ rotulo: sp3.Codigo, href: '#/sp/' + encodeURIComponent(sp3.Codigo) }] : [])
        .concat(a3 ? [{ rotulo: a3.Codigo, href: '#/a/' + encodeURIComponent(a3.Codigo) }] : [])
        .concat([{ rotulo: t.Codigo }])) +
      '<section class="ficha-hero" style="background:var(--cv-navy)">' +
      '<span class="eyebrow">Tarefa' + (a3 ? ' da atividade ' + esc(a3.Nome) : '') + '</span>' +
      '<h2>' + esc(t.Codigo) + ' — ' + esc(t.Nome) + '</h2>' +
      '<div class="meta">' + (t.Tipo_Tarefa ? '<span><i class="fas fa-gear" aria-hidden="true"></i> ' + esc(t.Tipo_Tarefa) + '</span>' : '') +
      (t.Responsavel ? '<span>· ' + esc(t.Responsavel) + '</span>' : '') +
      (t.Duracao_Estimada ? '<span>· Duração estimada: ' + esc(t.Duracao_Estimada) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da tarefa</h3><dl class="ficha-dl">' +
      campo('Descrição', t.Descricao && esc(t.Descricao), true) +
      campo('Tipo (CBOK 4.0)', t.Tipo_Tarefa && esc(t.Tipo_Tarefa)) +
      campo('Responsável', t.Responsavel && esc(t.Responsavel)) +
      campo('Sistema', t.Sistema ? chips(t.Sistema, 'fa-desktop') : null) +
      campo('Observações', t.Observacoes && esc(t.Observacoes), true) + '</dl></div>' +
      '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(t.Imagem_Bizagi, t.Nome) + '</div>' +
      secVinculos('Tarefa', cod) +
      '</div><aside>' +
      (a3 ? '<div class="pp-card"><h3><i class="fas fa-arrow-turn-up" aria-hidden="true"></i> Atividade (pai)</h3>' +
        '<a class="proc-card" href="#/a/' + encodeURIComponent(a3.Codigo) + '"><div class="topo"><div><span class="cod">' + esc(a3.Codigo) + '</span><div class="nome" style="font-size:var(--fs-sm)">' + esc(a3.Nome) + '</div></div></div></a></div>' : '') +
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
        return '<tr><td class="cod">' + esc(x.ID) + '</td><td><strong>' + tit + '</strong><div class="pp-muted" style="font-size:var(--fs-sm)">' +
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
  /* ── GRÁFICOS (SVG puro, sem dependências; cores do DS gov.br) ────── */
  var PAL = ['#222b54', '#005ca8', '#007d4e', '#155bcb', '#74c9ea', '#89bd2b', '#8a6d00', '#c5170b'];
  function svgWrap(titulo, conteudo, vb, altura) {
    return '<figure class="graf"><figcaption>' + esc(titulo) + '</figcaption>' +
      '<svg viewBox="' + vb + '" role="img" aria-label="' + esc(titulo) + '" ' +
      (altura ? 'style="height:' + altura + 'px"' : '') + '>' + conteudo + '</svg></figure>';
  }
  function grafDonut(titulo, dados) {           // [{rotulo, valor, cor}]
    var total = dados.reduce(function (a, b) { return a + b.valor; }, 0);
    if (!total) return svgWrap(titulo, '<text x="150" y="90" text-anchor="middle" font-size="12" fill="#5c5c5c">Sem dados</text>', '0 0 300 180');
    var ang = -Math.PI / 2, R = 62, r = 38, cx = 90, cy = 90, arcos = '';
    dados.forEach(function (d) {
      if (!d.valor) return;
      var a2 = ang + (d.valor / total) * Math.PI * 2, big = (a2 - ang) > Math.PI ? 1 : 0;
      var x1 = cx + R * Math.cos(ang), y1 = cy + R * Math.sin(ang), x2 = cx + R * Math.cos(a2), y2 = cy + R * Math.sin(a2);
      var x3 = cx + r * Math.cos(a2), y3 = cy + r * Math.sin(a2), x4 = cx + r * Math.cos(ang), y4 = cy + r * Math.sin(ang);
      arcos += '<path d="M' + x1.toFixed(1) + ' ' + y1.toFixed(1) + 'A' + R + ' ' + R + ' 0 ' + big + ' 1 ' + x2.toFixed(1) + ' ' + y2.toFixed(1) +
        'L' + x3.toFixed(1) + ' ' + y3.toFixed(1) + 'A' + r + ' ' + r + ' 0 ' + big + ' 0 ' + x4.toFixed(1) + ' ' + y4.toFixed(1) + 'Z" fill="' + d.cor + '">' +
        '<title>' + esc(d.rotulo) + ': ' + d.valor + ' (' + Math.round(d.valor / total * 100) + '%)</title></path>';
      ang = a2;
    });
    var leg = dados.map(function (d, i) {
      return '<g transform="translate(180,' + (34 + i * 22) + ')"><rect width="11" height="11" rx="2" fill="' + d.cor + '"></rect>' +
        '<text x="17" y="10" font-size="11" fill="#333">' + esc(d.rotulo) + ' (' + d.valor + ')</text></g>';
    }).join('');
    return svgWrap(titulo, arcos + '<text x="' + cx + '" y="' + (cy + 5) + '" text-anchor="middle" font-size="20" font-weight="700" fill="#222b54">' + total + '</text>' + leg, '0 0 380 180');
  }
  function grafBarras(titulo, dados, sufixo, meta) {   // [{rotulo, valor, cor}]
    if (!dados.length) return svgWrap(titulo, '', '0 0 480 120');
    var max = Math.max.apply(null, dados.map(function (d) { return d.valor; }).concat(meta ? [meta] : [1])) || 1;
    var lw = 132, bw = 320, h = 26, alt = dados.length * h + 24;
    var barras = dados.map(function (d, i) {
      var y = i * h + 8, w = Math.max(2, d.valor / max * bw);
      return '<text x="0" y="' + (y + 13) + '" font-size="11" fill="#333">' + esc(String(d.rotulo).slice(0, 22)) + '</text>' +
        '<rect x="' + lw + '" y="' + y + '" width="' + bw + '" height="15" rx="3" fill="#f0f0f0"></rect>' +
        '<rect x="' + lw + '" y="' + y + '" width="' + w.toFixed(1) + '" height="15" rx="3" fill="' + (d.cor || '#005ca8') + '">' +
        '<title>' + esc(d.rotulo) + ': ' + d.valor + (sufixo || '') + '</title></rect>' +
        '<text x="' + (lw + bw + 6) + '" y="' + (y + 12) + '" font-size="11" font-weight="700" fill="#222b54">' + d.valor + (sufixo || '') + '</text>';
    }).join('');
    var linhaMeta = meta ? '<line x1="' + (lw + meta / max * bw) + '" y1="2" x2="' + (lw + meta / max * bw) + '" y2="' + (alt - 16) +
      '" stroke="#c5170b" stroke-width="1.5" stroke-dasharray="4 3"><title>Meta: ' + meta + (sufixo || '') + '</title></line>' : '';
    return svgWrap(titulo, barras + linhaMeta, '0 0 500 ' + alt);
  }
  function grafFunil(titulo, etapas) {          // [{rotulo, valor}]
    var max = etapas[0] ? etapas[0].valor : 1, h = 30, alt = etapas.length * h + 12;
    var corpo = etapas.map(function (e, i) {
      var w = Math.max(6, (e.valor / (max || 1)) * 300), y = i * h + 4;
      return '<rect x="' + (150 - w / 2 + 150) + '" y="' + y + '" width="' + w.toFixed(1) + '" height="22" rx="3" fill="' + PAL[i % PAL.length] + '" opacity="0.9">' +
        '<title>' + esc(e.rotulo) + ': ' + e.valor + '</title></rect>' +
        '<text x="0" y="' + (y + 16) + '" font-size="11" fill="#333">' + esc(e.rotulo) + '</text>' +
        '<text x="' + (300 + 150 + w / 2 + 8) + '" y="' + (y + 16) + '" font-size="11" font-weight="700" fill="#222b54">' + e.valor + '</text>';
    }).join('');
    return svgWrap(titulo, corpo, '0 0 640 ' + alt);
  }
  function grafLinha(titulo, pontos) {          // [{rotulo, valor}]
    if (pontos.length < 2) return svgWrap(titulo, '<text x="200" y="70" text-anchor="middle" font-size="12" fill="#5c5c5c">Dados insuficientes</text>', '0 0 400 140');
    var w = 420, h = 120, pad = 30;
    var max = Math.max.apply(null, pontos.map(function (p) { return p.valor; })) || 1;
    var pts = pontos.map(function (p, i) {
      return [pad + i * ((w - pad * 2) / (pontos.length - 1)), h - 20 - (p.valor / max) * (h - 45)];
    });
    var linha = pts.map(function (p, i) { return (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1); }).join(' ');
    var area = linha + ' L' + pts[pts.length - 1][0].toFixed(1) + ' ' + (h - 20) + ' L' + pts[0][0].toFixed(1) + ' ' + (h - 20) + ' Z';
    var marcas = pts.map(function (p, i) {
      return '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="3.5" fill="#005ca8"><title>' + esc(pontos[i].rotulo) + ': ' + pontos[i].valor + '</title></circle>' +
        '<text x="' + p[0].toFixed(1) + '" y="' + (h - 6) + '" font-size="9" fill="#5c5c5c" text-anchor="middle">' + esc(pontos[i].rotulo) + '</text>';
    }).join('');
    return svgWrap(titulo, '<path d="' + area + '" fill="#005ca8" opacity="0.12"></path>' +
      '<path d="' + linha + '" fill="none" stroke="#005ca8" stroke-width="2.5"></path>' + marcas, '0 0 ' + w + ' ' + h);
  }
  function grafHeat(titulo, dados) {            // matriz risco compacta [{p,i,qtd}]
    var cel = 30, sz = 5 * cel, corpo = '';
    for (var i = 5; i >= 1; i--) {
      for (var pb = 1; pb <= 5; pb++) {
        var n = pb * i, cls = n >= 20 ? '#ffb1b1' : n >= 12 ? '#ffc2a1' : n >= 5 ? '#ffe396' : '#cdeccb';
        var achou = dados.filter(function (d) { return d.p === pb && d.i === i; })[0];
        var x = 26 + (pb - 1) * cel, y = (5 - i) * cel;
        corpo += '<rect x="' + x + '" y="' + y + '" width="' + (cel - 3) + '" height="' + (cel - 3) + '" rx="3" fill="' + cls + '"></rect>';
        if (achou) corpo += '<text x="' + (x + cel / 2 - 1.5) + '" y="' + (y + cel / 2 + 2) + '" font-size="12" font-weight="700" fill="#222b54" text-anchor="middle">' + achou.qtd + '<title>Prob. ' + pb + ' × Impacto ' + i + ': ' + achou.qtd + ' risco(s)</title></text>';
      }
      corpo += '<text x="18" y="' + ((5 - i) * cel + cel / 2) + '" font-size="10" fill="#5c5c5c" text-anchor="end">' + i + '</text>';
    }
    for (var k = 1; k <= 5; k++) corpo += '<text x="' + (26 + (k - 1) * cel + cel / 2 - 1.5) + '" y="' + (sz + 10) + '" font-size="10" fill="#5c5c5c" text-anchor="middle">' + k + '</text>';
    corpo += '<text x="88" y="' + (sz + 24) + '" font-size="10" fill="#5c5c5c">Probabilidade →</text>';
    return svgWrap(titulo, corpo, '0 0 200 ' + (sz + 30));
  }

  /* ── TELA: dashboard gerencial ────────────────────────────────────── */
  function renderDashboard() {
    var el = $('#viewDashboard');
    var procs = DADOS.procs, hoje = hojeISO();
    var porStatus = {};
    procs.forEach(function (p) { porStatus[p.Status_Mapeamento || 'Não iniciado'] = (porStatus[p.Status_Mapeamento || 'Não iniciado'] || 0) + 1; });
    var concl = porStatus['Concluído'] || 0, andam = porStatus['Em andamento'] || 0;
    var cobertura = procs.length ? Math.round(concl / procs.length * 100) : 0;
    var media = procs.length ? Math.round(procs.reduce(function (s, p) { return s + p.Percentual; }, 0) / procs.length) : 0;
    var atrasados = procs.filter(function (p) { return p.Prazo_Previsto && p.Prazo_Previsto < hoje && p._status !== 'concluido'; });
    var riscosAb = DADOS.riscos.filter(function (r) { return !/encerrad/i.test(String(r.Status || '')); });
    var criticos = riscosAb.filter(function (r) { return r._classe === 'Alto' || r._classe === 'Extremo'; });
    var indsMed = DADOS.inds.filter(function (x) { return x._sit && x._sit !== 'Sem medição'; });
    var indsOk = indsMed.filter(function (x) { return x._sit === 'Meta atingida'; });
    var pctInds = indsMed.length ? Math.round(indsOk.length / indsMed.length * 100) : 0;
    // marcos agregados (funil)
    var funil = MARCOS_ROTULOS.map(function (rot, i) {
      return { rotulo: 'M' + (i + 1) + ' · ' + rot, valor: procs.filter(function (p) { return simNao(p[MARCOS_CAMPOS[i]]); }).length };
    });
    // evolução por mês (conclusões acumuladas)
    var meses = {};
    procs.filter(function (p) { return p.Data_Conclusao; }).forEach(function (p) {
      var k = String(p.Data_Conclusao).slice(0, 7); meses[k] = (meses[k] || 0) + 1;
    });
    var acc = 0, linha = Object.keys(meses).sort().map(function (k) {
      acc += meses[k]; return { rotulo: k.slice(5) + '/' + k.slice(2, 4), valor: acc };
    });
    // heat de riscos
    var heat = {}; riscosAb.forEach(function (r) { var k = r.Probabilidade_1a5 + '|' + r.Impacto_1a5; heat[k] = (heat[k] || 0) + 1; });
    var heatArr = Object.keys(heat).map(function (k) { return { p: +k.split('|')[0], i: +k.split('|')[1], qtd: heat[k] }; });

    var CAT_COR = { gerencial: '#222b54', finalistico: '#005ca8', suporte: '#007d4e' };
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Dashboard gerencial</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp3)">Visão executiva da carteira de mapeamento — atualizada automaticamente a cada carga da planilha. Última leitura: <strong>' + esc(FONTE) + '</strong>.</p>' +
      '<div class="kpi-grid" style="margin-top:0">' +
      '<div class="kpi"><span class="num">' + cobertura + '%</span><span class="lbl">Cobertura da carteira</span><span class="sub">' + concl + ' de ' + procs.length + ' processos publicados</span></div>' +
      '<div class="kpi"><span class="num">' + media + '%</span><span class="lbl">Avanço médio</span><span class="sub">' + andam + ' em andamento</span></div>' +
      '<div class="kpi ' + (atrasados.length ? 'warn' : 'ok') + '"><span class="num">' + atrasados.length + '</span><span class="lbl">Prazos vencidos</span><span class="sub">sem conclusão registrada</span></div>' +
      '<div class="kpi ' + (criticos.length ? 'erro' : 'ok') + '"><span class="num">' + criticos.length + '</span><span class="lbl">Riscos críticos</span><span class="sub">' + riscosAb.length + ' riscos abertos no total</span></div>' +
      '<div class="kpi ' + (pctInds >= 70 ? 'ok' : 'warn') + '"><span class="num">' + pctInds + '%</span><span class="lbl">Indicadores na meta</span><span class="sub">' + indsOk.length + ' de ' + indsMed.length + ' medidos</span></div>' +
      '<div class="kpi"><span class="num">' + DADOS.tarefas.length + '</span><span class="lbl">Tarefas mapeadas</span><span class="sub">' + DADOS.ativs.length + ' atividades · ' + DADOS.subs.length + ' subprocessos</span></div>' +
      '</div>' +
      '<div class="graf-grid">' +
      grafDonut('Situação do mapeamento', [
        { rotulo: 'Concluído', valor: concl, cor: '#137436' },
        { rotulo: 'Em andamento', valor: andam, cor: '#8a6d00' },
        { rotulo: 'Não iniciado', valor: porStatus['Não iniciado'] || 0, cor: '#9e9e9e' },
        { rotulo: 'Suspenso', valor: porStatus['Suspenso'] || 0, cor: '#c5170b' }]) +
      grafDonut('Processos por tipo (CBOK)', ['gerencial', 'finalistico', 'suporte'].map(function (c) {
        return { rotulo: c === 'finalistico' ? 'Finalístico' : c === 'gerencial' ? 'Gerencial' : 'Suporte',
          valor: procs.filter(function (p) { var m = IDX.mp[p.Macroprocesso]; return m && m._cat === c; }).length, cor: CAT_COR[c] };
      })) +
      grafBarras('Avanço por macroprocesso (%)', DADOS.macros.map(function (m) {
        var ps = IDX.procsPorMacro[m.Codigo] || [];
        return { rotulo: m.Codigo + ' ' + m.Nome, cor: CAT_COR[m._cat] || '#005ca8',
          valor: ps.length ? Math.round(ps.reduce(function (s, p) { return s + p.Percentual; }, 0) / ps.length) : 0 };
      }), '%', 100) +
      grafFunil('Marcos concluídos na carteira (M1 → M9)', funil) +
      grafLinha('Processos publicados (acumulado)', linha) +
      grafHeat('Riscos abertos por probabilidade × impacto', heatArr) +
      grafBarras('Riscos abertos por categoria', (function () {
        var c = {}; riscosAb.forEach(function (r) { c[r.Categoria || '—'] = (c[r.Categoria || '—'] || 0) + 1; });
        return Object.keys(c).sort(function (a, b) { return c[b] - c[a]; }).map(function (k, i) { return { rotulo: k, valor: c[k], cor: PAL[i % PAL.length] }; });
      })()) +
      grafBarras('Documentos por situação', (function () {
        var c = {}; DADOS.docs.forEach(function (x) { c[x.Situacao || '—'] = (c[x.Situacao || '—'] || 0) + 1; });
        return Object.keys(c).map(function (k, i) { return { rotulo: k, valor: c[k], cor: PAL[i % PAL.length] }; });
      })()) +
      '</div>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Pontos de atenção</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> Processos com prazo vencido</h3>' +
      (atrasados.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Processo de negócio</th><th>Responsável</th><th>Prazo</th><th>Avanço</th></tr></thead><tbody>' +
        atrasados.map(function (p) {
          return '<tr data-link><td class="cod">' + esc(p.Codigo) + '</td><td><a href="#/p/' + encodeURIComponent(p.Codigo) + '"><strong>' + esc(p.Nome) + '</strong></a></td>' +
            '<td style="font-size:var(--fs-sm)">' + esc(p.Dono_Processo || '—') + '</td><td style="font-size:var(--fs-sm);white-space:nowrap">' + fmtData(p.Prazo_Previsto) + '</td>' +
            '<td style="min-width:120px">' + barraPct(p.Percentual) + '</td></tr>';
        }).join('') + '</tbody></table></div>' : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Nenhum processo com prazo vencido.</span></div></div>') + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Riscos críticos abertos</h3>' +
      (criticos.length ? tabelaRiscosHtml(criticos, true) : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Nenhum risco crítico em aberto.</span></div></div>') + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> Indicadores fora da meta</h3>' +
      (function () {
        var fora = DADOS.inds.filter(function (x) { return x._sit === 'Abaixo da meta' || x._sit === 'Acima da meta'; });
        return fora.length ? tabelaIndsHtml(fora, true) : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Todos os indicadores medidos estão na meta.</span></div></div>';
      })() + '</div></section>';
    $all('#viewDashboard tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        if (ev.target.closest('a')) return;
        var lk = tr.querySelector('a'); if (lk) location.hash = lk.getAttribute('href');
      });
    });
  }

  /* ── TELA: repositório (jornada + materiais + metodologia) ────────── */
  var filtroRepo = { cat: '', fase: '', q: '' };
  var FASES_JORNADA = ['Descobrir', 'Definir', 'Desenvolver', 'Entregar', 'Evoluir'];
  function cardRepo(it) {
    var interno = !/^https?:/i.test(String(it.Link || ''));
    var icone = { 'Documento oficial': 'fa-scale-balanced', 'Template': 'fa-file-lines',
      'Instrumento': 'fa-toolbox', 'Ferramenta': 'fa-screwdriver-wrench', 'Referência': 'fa-book' }[it.Categoria] || 'fa-file';
    return '<article class="repo-card"><div class="repo-topo">' +
      '<span class="repo-ico"><i class="fas ' + icone + '" aria-hidden="true"></i></span>' +
      '<div><span class="repo-cat">' + esc(it.Categoria || '') + (it.Fase_Ciclo ? ' · ' + esc(it.Fase_Ciclo) : '') + '</span>' +
      (it.Codigo ? '<span class="cod">' + esc(it.Codigo) + '</span>' : '') + '</div></div>' +
      '<h4>' + esc(it.Titulo) + '</h4><p>' + esc(it.Descricao || '') + '</p>' +
      '<div class="repo-rodape"><span class="repo-fonte">Fonte: ' + esc(it.Fonte || '—') + '</span>' +
      (it.Link ? '<a class="br-button secondary small" href="' + esc(it.Link) + '"' +
        (interno ? ' download' : ' target="_blank" rel="noopener"') + '>' +
        (interno ? '<i class="fas fa-download" aria-hidden="true"></i>&nbsp;Baixar' :
          '<i class="fas fa-up-right-from-square" aria-hidden="true"></i>&nbsp;Acessar<span class="sr-only"> (abre em nova aba)</span>') + '</a>' : '') +
      '</div></article>';
  }
  function renderRepositorio() {
    var el = $('#viewRepositorio');
    var repo = DADOS.repo;
    var cats = []; repo.forEach(function (i) { if (i.Categoria && cats.indexOf(i.Categoria) < 0) cats.push(i.Categoria); });
    var fases = []; repo.forEach(function (i) { if (i.Fase_Ciclo && fases.indexOf(i.Fase_Ciclo) < 0) fases.push(i.Fase_Ciclo); });
    var ql = filtroRepo.q.toLowerCase();
    var lista = repo.filter(function (i) {
      return (!filtroRepo.cat || i.Categoria === filtroRepo.cat) &&
        (!filtroRepo.fase || i.Fase_Ciclo === filtroRepo.fase) &&
        (!ql || String((i.Titulo || '') + ' ' + (i.Descricao || '') + ' ' + (i.Codigo || '')).toLowerCase().indexOf(ql) >= 0);
    });
    var met = DADOS.params.Link_Metodologia, guia = DADOS.params.Link_Guia;
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Repositório de materiais e ferramentas</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp3)">Tudo que apoia a gestão de processos na Codevasf: a jornada de mapeamento, a metodologia e o guia oficiais, os instrumentos por fase do ciclo BPM, os modelos e as ferramentas. O conteúdo desta aba vem das abas <strong>Jornada</strong> e <strong>Repositorio</strong> da planilha.</p>' +
      (met || guia ?
        '<div class="repo-oficial">' +
        (met ? '<a class="repo-oficial-card" href="' + esc(met) + '" target="_blank" rel="noopener"><i class="fas fa-scale-balanced" aria-hidden="true"></i><div><strong>Metodologia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicada na intranet/SEI</span></div><i class="fas fa-up-right-from-square seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        (guia ? '<a class="repo-oficial-card" href="' + esc(guia) + '" target="_blank" rel="noopener"><i class="fas fa-book-open" aria-hidden="true"></i><div><strong>Guia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicado na intranet/SEI</span></div><i class="fas fa-up-right-from-square seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        '</div>' : '') +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Jornada de mapeamento</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp2)">A jornada não começa no diagrama — começa na escuta. Das fases preliminares ao TO-BE e à melhoria contínua, com foco em quem executa, decide e se beneficia (design centrado no humano · Double Diamond · CBOK 4.0).</p>' +
      '<div class="jornada-fases" aria-hidden="true">' + FASES_JORNADA.map(function (f) { return '<span class="jf jf-' + slug(f) + '">' + f + '</span>'; }).join('<i class="fas fa-chevron-right"></i>') + '</div>' +
      '<ol class="jornada">' + DADOS.jornada.map(function (e) {
        return '<li class="jornada-etapa fase-' + slug(e.Fase || '') + '">' +
          '<div class="je-topo"><span class="je-num">' + esc(e.Ordem) + '</span><div><span class="je-fase">' + esc(e.Fase || '') + '</span><h4>' + esc(e.Nome) + '</h4></div><span class="je-dur"><i class="far fa-clock" aria-hidden="true"></i> ' + esc(e.Duracao || '') + '</span></div>' +
          '<p class="je-obj">' + esc(e.Objetivo || '') + '</p>' +
          '<div class="je-grid">' +
          '<div class="je-caixa"><b><i class="fas fa-list-check" aria-hidden="true"></i> Atividades-chave</b><ul>' + listar(e.Atividades_Chave).map(function (a) { return '<li>' + esc(a) + '</li>'; }).join('') + '</ul></div>' +
          '<div class="je-caixa"><b><i class="fas fa-people-group" aria-hidden="true"></i> Quem faz</b><p>' + esc(listar(e.Quem_Faz).join(' · ')) + '</p>' +
          '<b style="margin-top:8px"><i class="fas fa-box-open" aria-hidden="true"></i> Entregáveis</b><p>' + esc(listar(e.Entregaveis).join(' · ')) + '</p></div>' +
          '</div>' +
          (e.Sentimento_Usuario ? '<p class="je-sente"><i class="far fa-heart" aria-hidden="true"></i> ' + esc(e.Sentimento_Usuario) + '</p>' : '') +
          '</li>';
      }).join('') + '</ol></section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Instrumentos, modelos e ferramentas</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-filtros">' +
      '<select id="repoCat" aria-label="Filtrar por categoria"><option value="">Todas as categorias</option>' + cats.map(function (c) { return '<option' + (filtroRepo.cat === c ? ' selected' : '') + '>' + esc(c) + '</option>'; }).join('') + '</select>' +
      '<select id="repoFase" aria-label="Filtrar por fase do ciclo"><option value="">Todas as fases do ciclo</option>' + fases.map(function (c) { return '<option' + (filtroRepo.fase === c ? ' selected' : '') + '>' + esc(c) + '</option>'; }).join('') + '</select>' +
      '<input type="search" id="repoQ" placeholder="Buscar no repositório…" value="' + esc(filtroRepo.q) + '" aria-label="Buscar no repositório">' +
      '<span class="pp-muted" style="font-size:var(--fs-sm)">' + lista.length + ' de ' + repo.length + ' itens</span></div>' +
      (lista.length ? '<div class="repo-grid">' + lista.map(cardRepo).join('') + '</div>' : '<p class="pp-vazio">Nenhum item com esses filtros.</p>') + '</section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Metodologia em resumo</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-rotate" aria-hidden="true"></i> Ciclo de vida BPM (CBOK 4.0)</h3><ol class="ciclo">' +
      '<li><h4>Alinhamento à estratégia e metas</h4><p>Priorização da carteira e vínculo do processo aos objetivos institucionais.</p></li>' +
      '<li><h4>Arquitetar mudanças</h4><p>Modelagem (AS-IS), análise, desenho do estado futuro (TO-BE) e definição da medição.</p></li>' +
      '<li><h4>Desenvolver iniciativas</h4><p>Planos de implantação, capacitação, mudanças e tecnologia (visão PMBOK do projeto).</p></li>' +
      '<li><h4>Implementar mudanças</h4><p>Execução dos planos, publicação de procedimentos e estabilização.</p></li>' +
      '<li><h4>Medir o sucesso</h4><p>Monitoramento por indicadores e melhoria contínua (novo giro do ciclo).</p></li></ol></div>' +
      '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M9)</h3>' +
      '<p style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Roteiro-padrão de cada projeto de mapeamento, do primeiro contato com a área até a publicação no repositório:</p>' +
      '<ul class="marcos">' + MARCOS_ROTULOS.map(function (r) { return '<li class="feito"><span>' + esc(r) + '</span><i class="fas fa-check-circle" aria-hidden="true"></i></li>'; }).join('') + '</ul></div>' +
      '<div class="pp-card"><h3><i class="fas fa-database" aria-hidden="true"></i> Como este painel é alimentado</h3><dl class="ficha-dl">' +
      campo('1. Google Sheets (recomendado)', 'Importe a planilha para o Google Sheets, compartilhe como “qualquer pessoa com o link pode ver” e informe o ID em <code>PAINEL_CONFIG.googleSheetId</code> (index.html).', true) +
      campo('2. Planilha no repositório', 'Sem Google Sheets, o painel lê <code>data/painel-processos-dados.xlsx</code> publicado junto com o site.', true) +
      campo('3. Dados embutidos', 'Reserva (inclusive offline): <code>js/dados.js</code>, gerado por <code>scripts/planilha_para_js.py</code>.', true) +
      '</dl><div class="br-message warning" role="status" style="margin-top:var(--sp2)"><div class="icon"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i></div><div class="content"><span class="message-title">Dados fictícios.</span> <span class="message-body">Todo o conteúdo exibido foi criado apenas para demonstrar o painel — substitua na planilha.</span></div></div></div></section>';
    var f1 = $('#repoCat'), f2 = $('#repoFase'), f3 = $('#repoQ');
    if (f1) f1.onchange = function () { filtroRepo.cat = this.value; renderRepositorio(); };
    if (f2) f2.onchange = function () { filtroRepo.fase = this.value; renderRepositorio(); };
    if (f3) f3.oninput = function () { filtroRepo.q = this.value; renderRepositorio(); var n = $('#repoQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
  }

  /* ── TELA: NUGEP ──────────────────────────────────────────────────── */
  function iniciais(nome) {
    var p = String(nome || '').trim().split(/\s+/);
    return (((p[0] || '')[0] || '') + ((p.length > 1 ? p[p.length - 1][0] : '') || '')).toUpperCase();
  }
  function renderNugep() {
    var el = $('#viewNugep');
    var P = DADOS.params || {};
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>NUGEP — Núcleo de Gestão Normativa e de Processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="max-width:64rem;margin-bottom:var(--sp3)">Equipe multidisciplinar formada por integrantes de diferentes unidades da Codevasf, responsável pela condução do mapeamento, modelagem e melhoria contínua dos processos institucionais, em articulação com a Área de Estratégia e Finanças (AE), a Gerência de Planejamento Estratégico (GPE) e a Unidade de Gestão Normativa e de Processos (UNP), conforme a Metodologia e o Guia de Gerenciamento de Processos (RES 031/2025). Cadastre, altere ou remova integrantes na aba <strong>NUGEP</strong> da planilha.</p>' +
      (DADOS.nugep.length ? '<div class="nugep-grid">' + DADOS.nugep.map(function (m) {
        return '<article class="nugep-card"><span class="nugep-avatar" aria-hidden="true">' + esc(iniciais(m.Nome)) + '</span>' +
          '<h4>' + esc(m.Nome) + '</h4><p class="nugep-papel">' + esc(m.Papel || '') + '</p>' +
          '<p class="nugep-unid"><strong>' + esc(m.Unidade_Sigla || '') + '</strong>' + (m.Unidade_Nome ? '<br>' + esc(m.Unidade_Nome) : '') + '</p>' +
          (m.Email ? '<a href="mailto:' + esc(m.Email) + '"><i class="fas fa-envelope" aria-hidden="true"></i> ' + esc(m.Email) + '</a>' : '') +
          (m.Telefone ? '<a href="tel:+55' + esc(String(m.Telefone).replace(/\D/g, '')) + '"><i class="fas fa-phone" aria-hidden="true"></i> ' + esc(m.Telefone) + '</a>' : '') +
          '</article>';
      }).join('') + '</div>' : '<p class="pp-vazio">Nenhum integrante cadastrado na aba NUGEP da planilha.</p>') +
      '<div class="pp-card" style="margin-top:var(--sp4)"><h3><i class="fas fa-building" aria-hidden="true"></i> Contato institucional</h3>' +
      '<p style="font-size:var(--fs-sm)"><strong>' + esc(P.Contato_Unidade || 'Unidade de Gestão Normativa e de Processos (AE/GPE/UNP)') + '</strong><br>' +
      (P.Contato_Email ? 'E-mail: <a href="mailto:' + esc(P.Contato_Email) + '">' + esc(P.Contato_Email) + '</a>' : '') +
      (P.Contato_Telefone ? ' · Telefone: ' + esc(P.Contato_Telefone) : '') + '</p></div>';
  }

  /* ── TELA: glossário ──────────────────────────────────────────────── */
  var filtroGloss = { q: '', cat: '', letra: '' };
  function renderGlossario() {
    var el = $('#viewGlossario');
    var todos = DADOS.glossario;
    var cats = []; todos.forEach(function (t) { if (t.Categoria && cats.indexOf(t.Categoria) < 0) cats.push(t.Categoria); });
    var ql = filtroGloss.q.toLowerCase();
    var lista = todos.filter(function (t) {
      var letra = String(t.Termo || '').charAt(0).toUpperCase();
      return (!filtroGloss.cat || t.Categoria === filtroGloss.cat) &&
        (!filtroGloss.letra || letra === filtroGloss.letra) &&
        (!ql || String((t.Termo || '') + ' ' + (t.Definicao || '') + ' ' + (t.Termos_Relacionados || '')).toLowerCase().indexOf(ql) >= 0);
    });
    var letrasDisp = {}; todos.forEach(function (t) { letrasDisp[String(t.Termo || '').charAt(0).toUpperCase()] = 1; });
    var abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var porLetra = {};
    lista.forEach(function (t) { var L = String(t.Termo || '').charAt(0).toUpperCase(); (porLetra[L] = porLetra[L] || []).push(t); });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Glossário de Gestão de Processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp3)">' + todos.length + ' termos curados a partir do BPM CBOK 4.0 (ABPMP), do PMBOK (PMI), da ISO 31000 e da Metodologia de Gerenciamento de Processos da Codevasf (RES 031/2025) — mantidos na aba <strong>Glossario</strong> da planilha.</p>' +
      '<div class="pp-filtros"><input type="search" id="glossQ" placeholder="Buscar termo, sigla ou conceito (ex.: SIPOC, KPI, raia)…" value="' + esc(filtroGloss.q) + '" aria-label="Buscar termo" style="flex:1;min-width:230px">' +
      '<select id="glossCat" aria-label="Filtrar por categoria"><option value="">Todas as categorias (' + todos.length + ')</option>' +
      cats.map(function (c) { var n = todos.filter(function (t) { return t.Categoria === c; }).length; return '<option value="' + esc(c) + '"' + (filtroGloss.cat === c ? ' selected' : '') + '>' + esc(c) + ' (' + n + ')</option>'; }).join('') + '</select></div>' +
      '<div class="gloss-abc" role="group" aria-label="Filtrar por letra"><button type="button" class="' + (filtroGloss.letra ? '' : 'ativo') + '" data-letra="">Todos</button>' +
      abc.map(function (L) { return '<button type="button" data-letra="' + L + '" class="' + (filtroGloss.letra === L ? 'ativo' : '') + '"' + (letrasDisp[L] ? '' : ' disabled') + '>' + L + '</button>'; }).join('') + '</div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin:var(--sp2) 0">' + lista.length + ' resultado(s)</p>' +
      (lista.length ? Object.keys(porLetra).sort().map(function (L) {
        return '<h3 class="gloss-letra">' + L + '</h3><div class="gloss-grid">' + porLetra[L].map(function (t) {
          return '<article class="gloss-card"><div class="gloss-topo"><h4>' + esc(t.Termo) + '</h4><span class="gloss-cat">' + esc(t.Categoria || '') + '</span></div>' +
            '<p>' + esc(t.Definicao || '') + '</p>' +
            '<div class="gloss-rodape">' + (t.Fonte ? '<span class="repo-fonte">Fonte: ' + esc(t.Fonte) + '</span>' : '') +
            (t.Termos_Relacionados ? '<span class="chip-lista">' + listar(t.Termos_Relacionados).map(function (r) { return '<button type="button" class="chip gloss-rel" data-termo="' + esc(r) + '">' + esc(r) + '</button>'; }).join('') + '</span>' : '') +
            '</div></article>';
        }).join('') + '</div>';
      }).join('') : '<p class="pp-vazio">Nenhum termo encontrado com esses filtros.</p>');
    var q = $('#glossQ'), c = $('#glossCat');
    if (q) q.oninput = function () { filtroGloss.q = this.value; filtroGloss.letra = ''; renderGlossario(); var n = $('#glossQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
    if (c) c.onchange = function () { filtroGloss.cat = this.value; renderGlossario(); };
    $all('.gloss-abc button', el).forEach(function (b) { b.onclick = function () { filtroGloss.letra = b.getAttribute('data-letra'); renderGlossario(); }; });
    $all('.gloss-rel', el).forEach(function (b) { b.onclick = function () { filtroGloss.q = b.getAttribute('data-termo'); filtroGloss.letra = ''; filtroGloss.cat = ''; renderGlossario(); }; });
  }

  /* ── TELA: FAQ ────────────────────────────────────────────────────── */
  var filtroFaq = '';
  function renderFaq() {
    var el = $('#viewFaq');
    var todos = DADOS.faq;
    var cats = []; todos.forEach(function (f) { if (f.Categoria && cats.indexOf(f.Categoria) < 0) cats.push(f.Categoria); });
    var lista = filtroFaq ? todos.filter(function (f) { return f.Categoria === filtroFaq; }) : todos;
    var porCat = {}; lista.forEach(function (f) { (porCat[f.Categoria || 'Geral'] = porCat[f.Categoria || 'Geral'] || []).push(f); });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Perguntas frequentes</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp3)">Dúvidas sobre gestão DE e POR processos na Codevasf — CBOK 4.0, SIPOC, indicadores, riscos e o uso deste painel. Mantidas na aba <strong>FAQ</strong> da planilha.</p>' +
      '<div class="faq-cats"><button type="button" class="chip ' + (filtroFaq ? '' : 'ativo') + '" data-cat="">Todas (' + todos.length + ')</button>' +
      cats.map(function (c) { var n = todos.filter(function (f) { return f.Categoria === c; }).length; return '<button type="button" class="chip ' + (filtroFaq === c ? 'ativo' : '') + '" data-cat="' + esc(c) + '">' + esc(c) + ' (' + n + ')</button>'; }).join('') + '</div>' +
      Object.keys(porCat).map(function (cat) {
        return '<h3 class="faq-cat-h">' + esc(cat) + '</h3>' + porCat[cat].map(function (f) {
          return '<details class="faq-item"><summary><i class="fas fa-circle-question" aria-hidden="true"></i><span>' + esc(f.Pergunta) + '</span><i class="fas fa-chevron-down seta" aria-hidden="true"></i></summary><div class="faq-resp">' + esc(f.Resposta || '') + '</div></details>';
        }).join('');
      }).join('');
    $all('.faq-cats .chip', el).forEach(function (b) { b.onclick = function () { filtroFaq = b.getAttribute('data-cat'); renderFaq(); }; });
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
      t: DADOS.tarefas.filter(function (t) { return bate(t.Codigo) || bate(t.Nome) || bate(t.Descricao); }),
      doc: DADOS.docs.filter(function (x) { return bate(x.ID) || bate(x.Titulo); }),
      reg: DADOS.diario.filter(function (e) { return bate(e.Titulo) || bate(e.Descricao); }),
      gl: DADOS.glossario.filter(function (t) { return bate(t.Termo) || bate(t.Definicao); }),
      rp: DADOS.repo.filter(function (i) { return bate(i.Titulo) || bate(i.Descricao) || bate(i.Codigo); })
    };
    var total = r.mp.length + r.p.length + r.sp.length + r.a.length + r.doc.length + r.reg.length + r.gl.length + r.rp.length + r.t.length;
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
      grupo('Processos de negócio', r.p, function (p) { return linha('#/p/' + encodeURIComponent(p.Codigo), p.Codigo, p.Nome, esc(p.Status_Mapeamento) + ' · ' + p.Percentual + '%'); }) +
      grupo('Subprocessos', r.sp, function (s) { return linha('#/sp/' + encodeURIComponent(s.Codigo), s.Codigo, s.Nome, ''); }) +
      grupo('Atividades', r.a, function (a) { return linha('#/a/' + encodeURIComponent(a.Codigo), a.Codigo, a.Nome, esc(a.Responsavel_Ator || '')); }) +
      grupo('Tarefas', r.t, function (t) { return linha('#/t/' + encodeURIComponent(t.Codigo), t.Codigo, t.Nome, esc(t.Tipo_Tarefa || '')); }) +
      grupo('Documentos', r.doc, function (x) {
        return '<div class="doc-item"><i class="fas fa-file fa-stack-ico" aria-hidden="true"></i><div><div class="tit">' +
          (x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '</a>' : esc(x.Titulo)) +
          '</div><div class="meta">' + linkVinculo(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</div></div></div>';
      }) +
      grupo('Diário de mapeamento', r.reg, function (e) {
        return linha('#/p/' + encodeURIComponent(e.Processo), e.Processo, e.Titulo || '(registro)', fmtData(e.Data) + ' · ' + esc(e.Tipo || ''));
      }) +
      grupo('Glossário', r.gl, function (t) {
        return '<div class="doc-item"><i class="fas fa-spell-check fa-stack-ico" aria-hidden="true"></i><div><div class="tit"><a href="#/glossario">' + esc(t.Termo) + '</a></div><div class="meta">' + esc(String(t.Definicao || '').slice(0, 140)) + '…</div></div></div>';
      }) +
      grupo('Repositório de materiais', r.rp, function (i) {
        return '<div class="doc-item"><i class="fas fa-toolbox fa-stack-ico" aria-hidden="true"></i><div><div class="tit">' + (i.Link ? '<a href="' + esc(i.Link) + '" target="_blank" rel="noopener">' + esc(i.Titulo) + '</a>' : esc(i.Titulo)) + '</div><div class="meta">' + esc(i.Categoria || '') + '</div></div></div>';
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
    if ((c = $('#cntRepositorio'))) c.textContent = DADOS.repo.length;
    if ((c = $('#cntNugep'))) c.textContent = DADOS.nugep.length;
    if ((c = $('#cntGlossario'))) c.textContent = DADOS.glossario.length;
    if ((c = $('#cntFaq'))) c.textContent = DADOS.faq.length;
    montarAlertas();
    ligarAcoesCabecalho();
    if (window.PPUI) PPUI.setMenuSections([
      { rotulo: 'Início · Cadeia de Valor', href: '#/', icone: 'fa-house' },
      { rotulo: 'Catálogo de processos', href: '#/catalogo', icone: 'fa-layer-group' },
      { rotulo: 'Dashboard gerencial', href: '#/dashboard', icone: 'fa-chart-pie' },
      { rotulo: 'Repositório de materiais', href: '#/repositorio', icone: 'fa-toolbox' },
      { rotulo: 'Documentos', href: '#/documentos', icone: 'fa-folder-open' },
      { rotulo: 'Radar de riscos', href: '#/riscos', icone: 'fa-shield-halved' },
      { rotulo: 'Indicadores', href: '#/indicadores', icone: 'fa-chart-line' },
      { rotulo: 'Diário de mapeamento', href: '#/diario', icone: 'fa-timeline' },
      { rotulo: 'NUGEP', href: '#/nugep', icone: 'fa-people-group' },
      { rotulo: 'Glossário', href: '#/glossario', icone: 'fa-spell-check' },
      { rotulo: 'Perguntas frequentes', href: '#/faq', icone: 'fa-circle-question' }
    ]);
    rota();
  }
  function iniciar() {
    var v = $('#viewInicio');
    if (v) v.innerHTML = '<div class="pp-loading"><i class="fas fa-circle-notch" aria-hidden="true"></i> Carregando dados do painel…</div>';
    carregarDados().then(posCarga).catch(function (e) {
      console.error(e);
      if (v) v.innerHTML = '<div class="br-message warning" role="alert"><div class="icon"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i></div><div class="content"><span class="message-title">Não foi possível carregar os dados.</span> <span class="message-body">Verifique se data/painel-processos-dados.xlsx está publicado (ou gere js/dados.js com scripts/planilha_para_js.py). Detalhe: ' + esc(e.message) + '</span></div></div>';
    });
  }
  window.addEventListener('hashchange', rota);
  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
