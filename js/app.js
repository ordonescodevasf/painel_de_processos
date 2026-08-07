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
           'Documentos', 'Riscos', 'Indicadores',
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
    dd.jornada.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.repo.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.nugep.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.faq.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.glossario.sort(function (a, b) { return String(a.Termo || '').localeCompare(String(b.Termo || ''), 'pt-BR'); });
    dd.params = {};
    dd.parametros.forEach(function (x) { if (x.Chave) dd.params[x.Chave] = x.Valor || ''; });

    var idx = { mp: {}, p: {}, sp: {}, a: {}, t: {}, procsPorMacro: {}, subsPorPai: {}, ativsPorPai: {}, tarefasPorAtiv: {},
      vinc: { docs: {}, riscos: {}, inds: {} } };
    idx.ativsPorSub = idx.ativsPorPai;   // nome antigo do índice, mantido por compatibilidade
    dd.macros.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    // Codificação exibida dos macroprocessos, por tipo: gerencial = MG,
    // de suporte = MS, finalístico = MF, numerada de 1 em diante DENTRO de
    // cada tipo (MG-01, MG-02, MF-01…) — a contagem recomeça a cada tipo.
    // O código da planilha (MP-xx) segue sendo a chave de vínculo de
    // processos, documentos, riscos e indicadores: só a exibição muda.
    var seqCat = {};
    dd.macros.forEach(function (m) {
      var pre = m._cat === 'gerencial' ? 'MG' : m._cat === 'suporte' ? 'MS'
        : m._cat === 'finalistico' ? 'MF' : 'MP';
      seqCat[pre] = (seqCat[pre] || 0) + 1;
      m._cod = pre + '-' + ('0' + seqCat[pre]).slice(-2);
    });
    dd.macros.forEach(function (m) { idx.mp[m.Codigo] = m; });
    dd.procs.sort(function (a, b) { return String(a.Codigo).localeCompare(String(b.Codigo)); });
    dd.procs.forEach(function (p) {
      idx.p[p.Codigo] = p;
      (idx.procsPorMacro[p.Macroprocesso] = idx.procsPorMacro[p.Macroprocesso] || []).push(p);
    });
    dd.subs.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.subs.forEach(function (s) {
      idx.sp[s.Codigo] = s;
      // Vinculo_Pai aponta para um Processo (P-...) OU para outro Subprocesso (SP-...) —
      // o CBOK 4.0 não fixa a profundidade da decomposição ("Levels Vary in Number and
      // Name"): um subprocesso pode conter outro subprocesso, tantos níveis quanto o
      // processo exigir, até chegar à atividade.
      (idx.subsPorPai[s.Vinculo_Pai] = idx.subsPorPai[s.Vinculo_Pai] || []).push(s);
    });
    dd.ativs.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.ativs.forEach(function (a) {
      // Vinculo_Pai aceita o código de um Subprocesso (SP-...) OU de um Processo
      // (P-...): há processo que não tem subprocesso e se decompõe direto em
      // atividades (e estas em tarefas) — o CBOK 4.0 não obriga o nível
      // intermediário ("Levels Vary in Number and Name"). "Subprocesso" era o
      // nome da coluna até esta versão e segue aceito, para não quebrar
      // planilhas que ainda não foram regeradas.
      a._pai = String(a.Vinculo_Pai || a.Subprocesso || a.Processo || '').trim();
      idx.a[a.Codigo] = a;
      (idx.ativsPorPai[a._pai] = idx.ativsPorPai[a._pai] || []).push(a);
    });
    dd.tarefas.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.tarefas.forEach(function (t) {
      idx.t[t.Codigo] = t;
      (idx.tarefasPorAtiv[t.Atividade] = idx.tarefasPorAtiv[t.Atividade] || []).push(t);
    });
    function vincula(mapa, item) {
      listaVinculoPares(item.Vinculo_Nivel, item.Vinculo_Codigo).forEach(function (par) {
        var ch = par[0] + '|' + par[1];
        (mapa[ch] = mapa[ch] || []).push(item);
      });
    }
    dd.docs.forEach(function (x) { vincula(idx.vinc.docs, x); });
    dd.riscos.forEach(function (x) { vincula(idx.vinc.riscos, x); });
    dd.inds.forEach(function (x) { vincula(idx.vinc.inds, x); });
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
  // ── texto responsivo ──────────────────────────────────────────────
  // A mesma informação em 3 níveis de detalhe; o CSS mostra só a variante que
  // cabe na largura atual (breakpoints do gov.br DS: xs <576px · sm/md
  // 576–991.98px · lg ≥992px). As outras saem do fluxo com display:none, então
  // o leitor de tela lê uma única vez e a impressão usa sempre a completa.
  function txResp(curto, medio, completo) {
    return '<span class="pp-tx-resp">' +
      '<span class="tx-xs">' + curto + '</span>' +
      '<span class="tx-sm">' + medio + '</span>' +
      '<span class="tx-lg">' + completo + '</span></span>';
  }
  function plural(n, sing, plur) { return n + ' ' + (n === 1 ? sing : plur); }
  function campo(rotulo, valorHtml, span2, categoria) {
    var cls = (span2 ? 'span2 ' : '') + (categoria ? 'campo-' + categoria : '');
    return '<div' + (cls.trim() ? ' class="' + cls.trim() + '"' : '') + '><dt>' + esc(rotulo) + '</dt><dd>' +
      (valorHtml || '<span class="pp-vazio">Não informado</span>') + '</dd></div>';
  }
  // Sigla exibida de cada camada, sempre com DUAS letras, para que o nível
  // se leia no próprio código: MG/MF/MS (macroprocessos, por tipo), PP
  // (processo), SP (subprocesso), AT (atividade), TR (tarefa). O código da
  // planilha (MP-, P-, A-, T-) continua sendo a chave de vínculo — só a
  // exibição muda, então nenhuma planilha precisa ser regerada.
  function codDisp(c) {
    var s = String(c == null ? '' : c);
    if (IDX && IDX.mp[s]) return IDX.mp[s]._cod || s;
    if (s.indexOf('P-') === 0) return 'PP-' + s.slice(2);
    if (s.indexOf('A-') === 0) return 'AT-' + s.slice(2);
    if (s.indexOf('T-') === 0) return 'TR-' + s.slice(2);
    return s;                                // SP- e MG/MF/MS já têm 2 letras
  }
  var NIVEL_PREFIXO = { 'Macroprocesso': 'mp', 'Processo': 'p', 'Subprocesso': 'sp',
    'Atividade': 'a', 'Tarefa': 't' };
  var NIVEL_ROTULO = { 'Macroprocesso': 'Macroprocesso', 'Processo': 'Processo',
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
    return '<a href="' + rotaDe(nivel, codigo) + '"><span class="cod">' + esc(codDisp(codigo)) +
      '</span> ' + esc(nomeDe(nivel, codigo)) + '</a> <span class="pp-muted">(' + esc(nivelRotulo(nivel)) + ')</span>';
  }
  // Um mesmo item (documento, risco ou indicador) pode estar vinculado a dois ou mais
  // processos, dois ou mais subprocessos, ou a um processo e um subprocesso ao mesmo
  // tempo: Vinculo_Nivel e Vinculo_Codigo aceitam listas paralelas separadas por ';'.
  function listaVinculoPares(nivelStr, codigoStr) {
    var niveis = listar(nivelStr), codigos = listar(codigoStr), pares = [];
    for (var i = 0; i < Math.max(niveis.length, codigos.length); i++) {
      if (niveis[i] && codigos[i]) pares.push([niveis[i], codigos[i]]);
    }
    return pares;
  }
  function linkVinculos(nivelStr, codigoStr) {
    var pares = listaVinculoPares(nivelStr, codigoStr);
    if (!pares.length) return '<span class="pp-vazio">—</span>';
    return '<div class="vinculo-lista">' + pares.map(function (par) { return linkVinculo(par[0], par[1]); }).join('') + '</div>';
  }
  /* ── BREADCRUMB (br-breadcrumb, gov.br DS) ────────────────────────────
     Anatomia completa: 1 Botão (Home, terciário circular), 2 Separador
     (chevron-right decorativo), 3 Link (cada ancestral), 4 Título da
     página atual (texto comum, aria-current="page", não é link).
     Elementos opcionais também implementados: Truncamento a partir de 5
     itens, com Botão Terciário folder-plus abrindo os links intermediários
     em list dropdown, e tooltip (title) nos rótulos que podem truncar. */
  function breadcrumb(trilha) {   // [{rotulo, href?}]
    var itens = trilha.slice();
    var home = itens.shift() || { href: '#/' };
    var atual = itens.pop() || home;
    var oculto = [];
    // Nunca truncamos o Home, a última ancestral nem o título atual.
    if (itens.length + 2 > 5) {
      oculto = itens.slice(0, itens.length - 1);
      itens = itens.slice(-1);
    }
    var sep = '<i class="icon fas fa-chevron-right" aria-hidden="true"></i>';
    var html = '<nav class="br-breadcrumb" aria-label="Breadcrumbs">' +
      '<ol class="crumb-list" role="list">' +
      '<li class="crumb home"><a class="br-button circle" href="' + (home.href || '#/') +
      '" aria-label="Página inicial" title="Página inicial"><span class="sr-only">Página inicial</span>' +
      '<i class="fas fa-home" aria-hidden="true"></i></a></li>';
    if (oculto.length) {
      html += '<li class="crumb menu-mobil">' + sep +
        '<button class="br-button circle" type="button" data-crumb-menu aria-expanded="false"' +
        ' aria-label="Abrir menu flutuante" title="Abrir menu flutuante">' +
        '<span class="sr-only">Botão Menu</span><i class="fas fa-folder-plus" aria-hidden="true"></i>' +
        '</button></li>';
    }
    html += itens.map(function (t) {
      return '<li class="crumb">' + sep + (t.href
        ? '<a href="' + t.href + '" title="' + esc(t.rotulo) + '">' + esc(t.rotulo) + '</a>'
        : '<span title="' + esc(t.rotulo) + '">' + esc(t.rotulo) + '</span>') + '</li>';
    }).join('');
    html += '<li class="crumb" data-active="active">' + sep +
      '<span tabindex="0" aria-current="page" title="' + esc(atual.rotulo) + '">' +
      esc(atual.rotulo) + '</span></li></ol>';
    if (oculto.length) {
      html += '<nav class="br-card" aria-label="Níveis anteriores" hidden>' + oculto.map(function (t) {
        return '<div class="br-item"><a href="' + (t.href || '#/') + '" title="' + esc(t.rotulo) + '">' +
          esc(t.rotulo) + '</a></div>';
      }).join('') + '</nav>';
    }
    return html + '</nav>';
  }
  // Menu flutuante do truncamento: alterna folder-plus/folder-minus.
  document.addEventListener('click', function (ev) {
    var bt = ev.target.closest ? ev.target.closest('[data-crumb-menu]') : null;
    document.querySelectorAll('.br-breadcrumb > .br-card').forEach(function (card) {
      var dono = card.parentNode.querySelector('[data-crumb-menu]');
      var abrir = bt && dono === bt && card.hasAttribute('hidden');
      if (abrir) { card.removeAttribute('hidden'); } else { card.setAttribute('hidden', 'hidden'); }
      if (dono) {
        dono.setAttribute('aria-expanded', abrir ? 'true' : 'false');
        dono.setAttribute('aria-label', abrir ? 'Fechar menu flutuante' : 'Abrir menu flutuante');
        var ico = dono.querySelector('i');
        if (ico) ico.className = 'fas ' + (abrir ? 'fa-folder-minus' : 'fa-folder-plus');
      }
    });
  });
  /* ── Marcos do mapeamento (M1–M10) ──
     Nesta versão: "Reunião de contextualização" virou "Conhecer o processo";
     "Procedimento elaborado" saiu; "Processo publicado" (antes "Publicado no
     repositório") subiu para logo depois de "Procedimento aprovado"; e
     "Processo transformado" entrou como marco final.
     Cada marco lê a PRIMEIRA coluna existente da sua lista `campos` — o
     primeiro nome é o desta versão, os seguintes são os nomes antigos na
     planilha, mantidos para que planilhas ainda não regeradas continuem
     funcionando sem ajuste. */
  var MARCOS = [
    { rot: 'Conhecer o processo', campos: ['M1_Conhecer_Processo', 'M1_Reuniao_Contextualizacao'],
      desc: 'Primeira reunião com a área para conhecer o processo: apresentar a metodologia, entender o contexto e coletar o formulário de levantamento.' },
    { rot: 'Processo modelado', campos: ['M2_Macro_Processo_Modelados'],
      desc: 'Oficinas de modelagem do macroprocesso e do processo em BPMN, com a equipe de mapeamento e a área.' },
    { rot: 'Subprocessos modelados', campos: ['M3_Subprocessos_Modelados'],
      desc: 'Oficinas de modelagem dos subprocessos — podendo aprofundar vários níveis (subprocesso dentro de subprocesso) — inclusive identificando subprocessos ainda não mapeados.' },
    { rot: 'AS-IS modelado', campos: ['M4_ASIS_Modelado'],
      desc: 'O conjunto de diagramas AS-IS (macroprocesso, processo, subprocessos e atividades) está consolidado.' },
    { rot: 'AS-IS validado', campos: ['M5_ASIS_Validado'],
      desc: 'O dono do processo validou formalmente o AS-IS como retrato da realidade atual.' },
    { rot: 'Procedimento aprovado', campos: ['M6_Procedimento_Aprovado', 'M7_Procedimento_Aprovado'],
      desc: 'O procedimento (PRO) foi aprovado pela Diretoria Executiva (DEX) — pronto para orientar a execução do processo.' },
    { rot: 'Processo publicado', campos: ['M7_Processo_Publicado', 'M10_Publicado_Repositorio'],
      desc: 'O processo, seus diagramas e o PRO foram publicados no repositório institucional, disponíveis para consulta de toda a Empresa.' },
    { rot: 'TO-BE elaborado', campos: ['M8_TOBE_Elaborado'],
      desc: 'O redesenho (TO-BE) do processo foi elaborado, com melhorias, riscos residuais e indicadores propostos.' },
    { rot: 'TO-BE aprovado', campos: ['M9_TOBE_Aprovado', 'M9_TOBE_Validado'],
      desc: 'O dono do processo validou o TO-BE, que foi aprovado pela Diretoria Executiva (DEX).' },
    { rot: 'Processo transformado', campos: ['M10_Processo_Transformado'],
      desc: 'O redesenho aprovado foi implantado: o processo passou a ser executado na forma TO-BE, com os ganhos acompanhados pelos indicadores.' }
  ];
  var MARCOS_ROTULOS = MARCOS.map(function (m) { return m.rot; });
  var MARCOS_CAMPOS = MARCOS.map(function (m) { return m.campos[0]; });
  var MARCOS_DESCRICOES = MARCOS.map(function (m) { return m.desc; });
  // Valor do marco i no processo p — primeira coluna preenchida da lista.
  function valMarco(p, i) {
    var cs = (MARCOS[i] || {}).campos || [];
    for (var k = 0; k < cs.length; k++) {
      var v = p[cs[k]];
      if (v != null && String(v).trim() !== '') return v;
    }
    return '';
  }
  // Marco atual = o mais avançado já concluído (1 a 10; 0 = nenhum ainda).
  function marcoAtual(p) {
    var at = 0;
    for (var i = 0; i < MARCOS.length; i++) if (simNao(valMarco(p, i))) at = i + 1;
    return at;
  }
  function marcoRotulo(n) {
    return n ? 'M' + n + ' - ' + MARCOS_ROTULOS[n - 1] : 'Nenhum marco concluído';
  }
  // Estado de um marco: concluído, "não se aplica" ou pendente. O terceiro
  // estado existe para casos legítimos como o M3 (Subprocessos modelados) num
  // processo que não tem subprocessos — sem ele, o marco ficaria eternamente
  // pendente e o processo nunca pareceria concluído.
  function marcoEstado(v) {
    var s = String(v == null ? '' : v).trim();
    if (v === true || /^s/i.test(s)) return 'feito';
    if (/^n(ã|a)o\s*se\s*aplica$/i.test(s) || /^n\/?a$/i.test(s)) return 'na';
    return '';
  }
  /* Marcos do mapeamento pelo Componente Step (tipo complexo, indicador
     numérico, orientação horizontal com data-scroll): a jornada tem ordem
     lógica linear, que é exatamente o caso de uso do componente. Os
     indicadores não são interativos (o painel só informa o progresso),
     por isso vêm com disabled; a etapa atual é o primeiro marco ainda
     pendente e "não se aplica" recebe alerta em cor de atenção. */
  function marcosHtml(p) {
    var est = MARCOS.map(function (mk, i) { return marcoEstado(valMarco(p, i)); });
    var atual = -1;
    for (var k = 0; k < est.length; k++) { if (est[k] === '') { atual = k; break; } }
    var total = MARCOS.length;
    return '<nav class="br-step" data-label="bottom" data-scroll="data-scroll" role="none">' +
      '<div class="step-progress" role="listbox" aria-orientation="horizontal" aria-label="Marcos do mapeamento">' +
      MARCOS.map(function (mk, i) {
        var e = est[i], feito = e === 'feito', na = e === 'na', ativo = i === atual;
        var rotulo = 'M' + (i + 1) + ' — ' + MARCOS_ROTULOS[i] +
          (feito ? ' (concluído)' : na ? ' (não se aplica)' : ativo ? ' (em andamento)' : ' (pendente)');
        return '<button class="step-progress-btn' + (feito ? ' is-done' : '') + '" type="button" role="option"' +
          ' step-num="' + (i + 1) + '" aria-posinset="' + (i + 1) + '" aria-setsize="' + total + '"' +
          ' aria-selected="' + (ativo ? 'true' : 'false') + '"' + (ativo ? ' active' : '') + ' disabled' +
          (na ? ' data-alert="warning"' : '') +
          ' aria-label="' + esc(rotulo) + '" title="' + esc(MARCOS_DESCRICOES[i] + (na ? ' — Não se aplica a este processo.' : '')) + '">' +
          '<span class="step-info">' + esc(MARCOS_ROTULOS[i]) + '</span>' +
          (feito ? '<i class="step-icon fas fa-check" aria-hidden="true"></i>' : '') +
          (na ? '<span class="step-alert"></span>' : '') +
          '</button>';
      }).join('') +
      '</div></nav>';
  }
  // Sobe a cadeia de um subprocesso até achar seu Processo — como o CBOK permite
  // subprocesso dentro de subprocesso (profundidade variável), o "pai" de um
  // subprocesso pode ser outro subprocesso (código "SP-...") em vez de um
  // processo (código "P-...") direto. Retorna [maisFundo, ..., maisRaso],
  // sempre terminando no subprocesso que é filho direto do Processo.
  function cadeiaSubprocessos(spCodigo) {
    var cadeia = [], atual = IDX.sp[spCodigo], guarda = {};
    while (atual && !guarda[atual.Codigo]) {
      cadeia.push(atual); guarda[atual.Codigo] = true;
      if (String(atual.Vinculo_Pai || '').indexOf('SP-') === 0) atual = IDX.sp[atual.Vinculo_Pai];
      else break;
    }
    return cadeia;
  }
  function processoDoSubprocesso(spCodigo) {
    var cadeia = cadeiaSubprocessos(spCodigo);
    var raiz = cadeia[cadeia.length - 1];
    return raiz ? IDX.p[raiz.Vinculo_Pai] : null;
  }
  function ehCodigoSub(c) { return String(c || '').indexOf('SP-') === 0; }
  function atividadesDe(codigo) { return IDX.ativsPorPai[codigo] || []; }
  // Pai direto de uma atividade: um Subprocesso (SP-...) ou o próprio Processo
  // (P-...), quando o processo não tem subprocessos.
  function paiDaAtividade(a) {
    if (!a) return null;
    var c = a._pai || a.Vinculo_Pai || a.Subprocesso || '';
    if (ehCodigoSub(c)) return IDX.sp[c] ? { tipo: 'sp', item: IDX.sp[c] } : null;
    return IDX.p[c] ? { tipo: 'p', item: IDX.p[c] } : null;
  }
  // Ancestrais de uma atividade, prontos para o breadcrumb e para o card
  // "Navegar para": macroprocesso, processo e a cadeia de subprocessos — que
  // vem vazia quando a atividade pende direto do processo.
  function ancestraisDaAtividade(a) {
    var pai = paiDaAtividade(a), sp = null, p = null;
    if (pai && pai.tipo === 'sp') { sp = pai.item; p = processoDoSubprocesso(sp.Codigo); }
    else if (pai) { p = pai.item; }
    return { pai: pai, sp: sp, p: p, mp: p ? IDX.mp[p.Macroprocesso] : null,
      cadeiaSp: sp ? cadeiaSubprocessos(sp.Codigo).slice().reverse() : [] };
  }
  function contarAtividadesRecursivo(codigoPai) {
    // atividades ligadas DIRETO a codigoPai (processo sem subprocesso, ou
    // subprocesso) + as de todos os subprocessos descendentes, em qualquer
    // profundidade (subprocesso dentro de subprocesso)
    var total = atividadesDe(codigoPai).length;
    (IDX.subsPorPai[codigoPai] || []).forEach(function (s) {
      total += contarAtividadesRecursivo(s.Codigo);
    });
    return total;
  }
  function contarSubprocessosRecursivo(codigoPai) {
    var subs = IDX.subsPorPai[codigoPai] || [], total = subs.length;
    subs.forEach(function (s) { total += contarSubprocessosRecursivo(s.Codigo); });
    return total;
  }
  function contarTarefasRecursivo(codigoPai) {
    var total = 0;
    atividadesDe(codigoPai).forEach(function (a) { total += (IDX.tarefasPorAtiv[a.Codigo] || []).length; });
    (IDX.subsPorPai[codigoPai] || []).forEach(function (s) { total += contarTarefasRecursivo(s.Codigo); });
    return total;
  }
  var HIER_INFO = {
    mp: { rotulo: 'Macroprocesso', icone: 'fa-diagram-project', classe: 'hier-mp' },
    p: { rotulo: 'Processo', icone: 'fa-briefcase', classe: 'hier-p' },
    sp: { rotulo: 'Subprocesso', icone: 'fa-sitemap', classe: 'hier-sp' },
    a: { rotulo: 'Atividade', icone: 'fa-list-check', classe: 'hier-a' },
    t: { rotulo: 'Tarefa', icone: 'fa-check', classe: 'hier-t' }
  };
  // Card único de navegação hierárquica — reúne o acesso a TODOS os
  // ancestrais (do Macroprocesso até o pai mais próximo) num só lugar,
  // em vez de vários cards avulsos repetindo "suba um nível". Cada item:
  // { tipo: 'mp'|'p'|'sp'|'a', codigo, nome, href }.
  function cardHierarquia(itens) {
    if (!itens.length) return '';
    return '<div class="pp-card"><h3><i class="fas fa-route" aria-hidden="true"></i> Navegar para</h3><div class="hier-lista">' +
      itens.map(function (it) {
        var info = HIER_INFO[it.tipo];
        var catCls = it.tipo === 'mp' && it.cat ? ' cat-' + it.cat : '';
        return '<a class="hier-item ' + info.classe + catCls + '" href="' + it.href + '">' +
          '<span class="hier-ic"><i class="fas ' + info.icone + '" aria-hidden="true"></i></span>' +
          '<span class="hier-tx"><span class="hier-tipo">' + info.rotulo + '</span>' +
          '<span class="hier-nome">' + esc(codDisp(it.codigo)) + ' — ' + esc(it.nome) + '</span></span>' +
          '<i class="fas fa-chevron-right" aria-hidden="true"></i></a>';
      }).join('') + '</div></div>';
  }
  function urlDrive(u) {                     // link de compartilhamento → imagem exibível
    var m = String(u || '').match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]{20,})/);
    return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w2000' : null;
  }
  function diagramaHtml(caminho, titulo) {
    if (!caminho) return '<p class="pp-vazio">Diagrama ainda não publicado para este item. No Bizagi Modeler, use Publish → Web, suba o pacote num servidor e cole a URL na coluna Imagem_Bizagi da planilha.</p>';
    var href = esc(caminho);
    // Diagrama incorporado por iframe (Bizagi Web Publish) — mesmo padrão já
    // usado internamente na Codevasf (Base de Conhecimento / Wiki.js, AA/GTI)
    // para o Gerenciamento de Incidentes de TI. Exige que o servidor que
    // hospeda a publicação permita ser incorporado por outra origem (o
    // painel roda no GitHub Pages): Content-Security-Policy: frame-ancestors
    // precisa incluir o domínio do painel. Sem essa liberação, o quadro
    // abaixo aparece em branco — daí o link de abrir direto, que sempre
    // funciona independente disso.
    return '<div class="diagrama-frame diagrama-iframe">' +
      '<iframe src="' + href + '" title="Diagrama BPMN interativo — ' + esc(titulo) + '" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>' +
      '</div>' +
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-top:6px"><i class="fas fa-circle-info" aria-hidden="true"></i> Diagrama interativo (Bizagi Web Publish). Se o quadro acima aparecer em branco, o servidor pode estar bloqueando a incorporação ou exigir rede interna — use o link abaixo.</p>' +
      '<div class="diagrama-acoes"><a class="br-button secondary small" href="' + href +
      '" target="_blank" rel="noopener"><i class="fas fa-up-right-from-square" aria-hidden="true"></i>&nbsp;Abrir em nova aba<span class="sr-only"> (abre em nova aba)</span></a></div>';
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
  // Tabela de atividades — usada na ficha do Subprocesso e na ficha do
  // Processo (quando as atividades penduram direto no processo).
  function tabelaAtividadesHtml(ativs, vazio) {
    if (!ativs.length) return '<p class="pp-vazio">' + (vazio || 'Nenhuma atividade cadastrada.') + '</p>';
    return '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>#</th><th>Atividade</th><th>Responsável (ator)</th><th>Entradas</th><th>Saídas</th><th>Prazo</th></tr></thead><tbody>' +
      ativs.map(function (a, i) {
        var nt = (IDX.tarefasPorAtiv[a.Codigo] || []).length;
        return '<tr data-link><td>' + (i + 1) + '</td><td><a href="#/a/' + encodeURIComponent(a.Codigo) + '"><strong>' + esc(a.Nome) + '</strong></a>' +
          '<div class="cod">' + esc(codDisp(a.Codigo)) + (nt ? ' · ' + plural(nt, 'tarefa', 'tarefas') : '') + '</div></td>' +
          '<td style="font-size:var(--fs-sm)">' + esc(a.Responsavel_Ator || '—') + '</td>' +
          '<td style="font-size:var(--fs-sm)">' + (listar(a.Entradas).map(esc).join('; ') || '—') + '</td>' +
          '<td style="font-size:var(--fs-sm)">' + (listar(a.Saidas).map(esc).join('; ') || '—') + '</td>' +
          '<td style="font-size:var(--fs-sm);white-space:nowrap">' + esc(a.Prazo_Padrao || '—') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }
  function ligarLinhasTabela() {
    $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        if (ev.target.closest('a')) return;
        var lk = tr.querySelector('a'); if (lk) location.hash = lk.getAttribute('href');
      });
    });
  }
  function tabelaRiscosHtml(riscos, comVinculo) {
    if (!riscos.length) return '<p class="pp-vazio">Nenhum risco registrado.</p>';
    return '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>ID</th>' +
      (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Risco</th><th>P</th><th>I</th><th>P×I</th><th>Nível</th><th>Resposta</th><th>Status</th></tr></thead><tbody>' +
      riscos.map(function (r) {
        return '<tr id="risco-' + esc(r.ID) + '"><td class="cod">' + esc(r.ID) + '</td>' +
          (comVinculo ? '<td>' + linkVinculos(r.Vinculo_Nivel, r.Vinculo_Codigo) + '</td>' : '') +
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
      '<th>Indicador</th>' + (comVinculo ? '<th>Vinculado a</th>' : '') +
      '<th>Meta</th><th>Resultado</th><th>Situação</th><th>Periodicidade</th><th>Última medição</th></tr></thead><tbody>' +
      inds.map(function (x) {
        var cls = x._sit === 'Meta atingida' ? 'sit-ok' : (x._sit === 'Sem medição' || x._sit === 'Sem meta') ? 'sit-neutra' : 'sit-ruim';
        var un = x.Unidade ? ' ' + esc(x.Unidade) : '';
        return '<tr><td class="cod">' + esc(x.ID) + '</td>' +
          '<td><strong>' + esc(x.Nome) + '</strong>' +
          (x.Descricao_Formula ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(x.Descricao_Formula) + '</div>' : '') +
          '</td>' + (comVinculo ? '<td>' + linkVinculos(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' : '') +
          '<td>' + (x.Meta != null ? x.Meta + un : '—') + '</td><td>' + (x.Resultado_Atual != null ? x.Resultado_Atual + un : '—') +
          '</td><td class="' + cls + '">' + esc(x._sit) + '</td><td>' + esc(x.Periodicidade || '—') +
          '</td><td>' + fmtData(x.Ultima_Medicao) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }

  /* ── PAGINATION (br-pagination, gov.br DS) ────────────────────────────
     Montado com os seis módulos do componente:
       1 Setas de Navegação      Button (circular)
       2 Identificadores de Páginas  Button
       3 Botão Reticências       Button
       4 Módulo de Exibição      Select — quantos itens por página
       5 Módulo de Informação    Tipografia — faixa exibida e total
       6 Módulo de Atalho        Select — pular direto para uma página
     Só as Setas de Navegação são obrigatórias no DS; os outros quatro
     módulos são opcionais e estão todos aqui. Aparece em toda lista com
     mais de um item; some quando a lista tem zero ou um. ── */
  var PAG = {};
  function pagEstado(chave, porPag) {
    if (!PAG[chave]) PAG[chave] = { pag: 1, porPag: porPag };
    return PAG[chave];
  }
  function pagTotal(chave, n) {
    var e = PAG[chave];
    return Math.max(1, Math.ceil(n / e.porPag));
  }
  // Recorta a lista na página atual, corrigindo a página quando um filtro
  // encurta a lista e a página corrente deixa de existir.
  function pagFatia(chave, lista, porPag) {
    var e = pagEstado(chave, porPag), tot = pagTotal(chave, lista.length);
    if (e.pag > tot) e.pag = tot;
    if (e.pag < 1) e.pag = 1;
    return lista.slice((e.pag - 1) * e.porPag, e.pag * e.porPag);
  }
  // Janela de páginas: primeira, última, a atual e as vizinhas; o resto
  // vira Botão Reticências (módulo 3), que abre em dropdown as páginas
  // ocultas — por isso cada lacuna carrega a lista que representa.
  function pagJanela(atual, total) {
    var t = [], i;
    if (total <= 7) { for (i = 1; i <= total; i++) t.push(i); return t; }
    function faixa(a, b) { var r = []; for (var k = a; k <= b; k++) r.push(k); return r; }
    var ini = Math.max(2, atual - 1), fim = Math.min(total - 1, atual + 1);
    t.push(1);
    if (ini > 2) t.push({ ocultas: faixa(2, ini - 1) });
    for (i = ini; i <= fim; i++) t.push(i);
    if (fim < total - 1) t.push({ ocultas: faixa(fim + 1, total - 1) });
    t.push(total);
    return t;
  }
  // Módulos de opção (Exibir e Página) usam o Select personalizado do DS:
  // campo em densidade alta sem contorno + lista de opções em br-radio.
  function pagSelectHtml(id, rotulo, ops, valor, marca) {
    return '<div class="br-select" data-pag-select><div class="br-input">' +
      '<label for="' + id + '">' + rotulo + '</label>' +
      '<input id="' + id + '" type="text" readonly value="' + valor + '" ' + marca +
      ' role="combobox" aria-expanded="false" aria-controls="' + id + '-lista">' +
      '<button class="br-button" type="button" tabindex="-1" data-trigger aria-label="Exibir lista">' +
      '<i class="fas fa-angle-down" aria-hidden="true"></i></button></div>' +
      '<div class="br-list" id="' + id + '-lista" role="listbox" tabindex="0">' +
      ops.map(function (n) {
        var oid = id + '-' + n, sel = n === valor;
        return '<div class="br-item' + (sel ? ' selected' : '') + '" tabindex="-1" role="option" aria-selected="' +
          sel + '"><div class="br-radio"><input id="' + oid + '" type="radio" name="' + id + '-op" value="' + n + '"' +
          (sel ? ' checked' : '') + '><label for="' + oid + '">' + n + '</label></div></div>';
      }).join('') + '</div></div>';
  }
  function paginacaoHtml(chave, total, rotulo, opcoes) {
    if (total <= 1) return '';
    var e = PAG[chave], tot = pagTotal(chave, total);
    var de = (e.pag - 1) * e.porPag + 1, ate = Math.min(total, e.pag * e.porPag);
    var ops = opcoes || [5, 10, 20, 50];
    var pgs = pagJanela(e.pag, tot).map(function (n) {
      if (typeof n === 'object') {
        return '<li class="pagination-ellipsis"><button type="button" class="br-button circle" data-pag-reticencias' +
          ' aria-expanded="false" aria-haspopup="true" aria-label="Abrir ou fechar a lista de páginas ocultas">' +
          '<i class="fas fa-ellipsis-h" aria-hidden="true"></i></button><div class="br-list" role="menu" hidden>' +
          n.ocultas.map(function (p) {
            return '<button type="button" class="br-item" role="menuitem" data-pag-ir="' + p +
              '" aria-label="Ir para a página ' + p + '">' + p + '</button>';
          }).join('') + '</div></li>';
      }
      return '<li><button type="button" class="page" data-pag-ir="' + n + '" aria-label="Ir para a página ' + n + '"' +
        (n === e.pag ? ' aria-current="page"' : '') + '>' + n + '</button></li>';
    }).join('');
    return '<nav class="br-pagination" data-pag="' + esc(chave) + '" aria-label="Paginação — ' + esc(rotulo) + '"' +
      ' data-total="' + tot + '" data-current="' + e.pag + '" data-per-page="' + e.porPag + '">' +
      '<div class="pagination-per-page">' +
      pagSelectHtml('pag-' + esc(chave) + '-exibir', 'Exibir', ops, e.porPag, 'data-pag-por') + '</div>' +
      '<hr class="vertical-sep">' +
      '<span class="pagination-information"><span class="current">' + de + '–' + ate +
      '</span> de <span class="total">' + total + '</span> ' + esc(rotulo) + '</span>' +
      '<ul class="pagination-pages">' + pgs + '</ul>' +
      '<hr class="vertical-sep">' +
      '<div class="pagination-go-to-page">' +
      pagSelectHtml('pag-' + esc(chave) + '-ir', 'Página',
        (function () { var o = []; for (var i = 1; i <= tot; i++) o.push(i); return o; })(), e.pag, 'data-pag-atalho') +
      '</div>' +
      '<div class="pagination-arrows">' +
      '<button type="button" class="br-button circle" data-previous-page data-pag-passo="-1"' + (e.pag === 1 ? ' disabled' : '') +
      ' aria-label="Voltar página"><i class="fas fa-angle-left" aria-hidden="true"></i></button>' +
      '<button type="button" class="br-button circle" data-next-page data-pag-passo="1"' + (e.pag === tot ? ' disabled' : '') +
      ' aria-label="Página seguinte"><i class="fas fa-angle-right" aria-hidden="true"></i></button>' +
      '</div></nav>';
  }
  // Liga os seis módulos de todas as paginações dentro de um container.
  function ligarPaginacao(escopo, redesenhar) {
    $all('.br-pagination', escopo).forEach(function (nav) {
      var chave = nav.getAttribute('data-pag'), e = PAG[chave];
      if (!e) return;
      function vai(n) { e.pag = n; redesenhar(); }
      $all('[data-pag-ir]', nav).forEach(function (b) {
        b.onclick = function () { vai(+b.getAttribute('data-pag-ir')); };
      });
      $all('[data-pag-passo]', nav).forEach(function (b) {
        b.onclick = function () { vai(e.pag + (+b.getAttribute('data-pag-passo'))); };
      });
      var sel = nav.querySelector('[data-pag-por]');
      var atalho = nav.querySelector('[data-pag-atalho]');
      // Select personalizado: abre/fecha a lista, aplica a opção escolhida
      // e fecha com Esc (recomendação de acessibilidade do componente).
      $all('[data-pag-select]', nav).forEach(function (bs) {
        var campo = bs.querySelector('input[type="text"]');
        var lista = bs.querySelector('.br-list');
        function abrir(v) {
          if (v) lista.setAttribute('expanded', ''); else lista.removeAttribute('expanded');
          campo.setAttribute('aria-expanded', String(!!v));
        }
        bs.querySelector('[data-trigger]').onclick = function () { abrir(!lista.hasAttribute('expanded')); };
        campo.onclick = function () { abrir(!lista.hasAttribute('expanded')); };
        campo.onkeydown = function (ev) {
          if (ev.key === 'Escape') abrir(false);
          if (ev.key === 'ArrowDown') { abrir(true); ev.preventDefault(); }
        };
        lista.onkeydown = function (ev) { if (ev.key === 'Escape') { abrir(false); campo.focus(); } };
        $all('.br-radio input', lista).forEach(function (r) {
          r.onchange = function () {
            abrir(false);
            if (campo === sel) { e.porPag = +r.value; e.pag = 1; redesenhar(); }
            else if (campo === atalho) { vai(+r.value); }
          };
        });
      });
      // Botão Reticências: dropdown com as páginas ocultas.
      $all('[data-pag-reticencias]', nav).forEach(function (b) {
        var lista = b.nextElementSibling;
        b.onclick = function () {
          var aberto = !lista.hidden;
          $all('.pagination-ellipsis .br-list', nav).forEach(function (l) { l.hidden = true; });
          $all('[data-pag-reticencias]', nav).forEach(function (o) { o.setAttribute('aria-expanded', 'false'); });
          lista.hidden = aberto;
          b.setAttribute('aria-expanded', String(!aberto));
        };
        lista.onkeydown = function (ev) { if (ev.key === 'Escape') { lista.hidden = true; b.setAttribute('aria-expanded', 'false'); b.focus(); } };
      });
    });
  }

  /* ── roteador + abas ──────────────────────────────────────────────── */
  var ROTAS_ABA = { inicio: '#/', catalogo: '#/catalogo', dashboard: '#/dashboard', repositorio: '#/repositorio',
    documentos: '#/documentos', riscos: '#/riscos', indicadores: '#/indicadores',
    nugep: '#/nugep', glossario: '#/glossario', faq: '#/faq' };
  var MAIS_ITENS = { repositorio: 'Repositório', nugep: 'NUGEP',
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


  /* ── ações do cabeçalho ───────────────────────────────────────────── */
  window.refreshAll = function () {
    var chip = $('#syncChip'); if (chip) chip.textContent = 'Atualizando…';
    carregarDados().then(function () { posCarga(); }).catch(function (e) {
      if (chip) chip.textContent = 'Falha ao atualizar';
      console.error(e);
    });
  };
  // Exporta uma matriz de linhas como CSV (BOM + ponto e vírgula, para
   // abrir direto no Excel em pt-BR).
  function baixarCsv(nome, matriz) {
    var txt = matriz.map(function (l) {
      return l.map(function (c) { return '"' + String(c == null ? '' : c).replace(/"/g, '""') + '"'; }).join(';');
    }).join('\r\n');
    var blob = new Blob(['\ufeff' + txt], { type: 'text/csv;charset=utf-8' });
    var a = d.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nome;
    d.body.appendChild(a); a.click(); a.remove();
  }
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
          '<span class="cv-tx"><span class="cod">' + esc(m._cod || m.Codigo) + '</span><span class="nome">' + esc(m.Nome) +
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
    var ger = DADOS.macros.filter(function (m) { return m._cat === 'gerencial'; });
    var fin = DADOS.macros.filter(function (m) { return m._cat === 'finalistico'; });
    var sup = DADOS.macros.filter(function (m) { return m._cat === 'suporte'; });
    el.innerHTML =
      '<section class="pp-hero">' +
      '<h1>Mapeamento de processos da Codevasf</h1>' +
      '</section>' +
      '<div class="kpi-grid">' +
      '<div class="kpi" title="Total de macroprocessos na cadeia de valor institucional (gerenciais, finalísticos e de suporte)."><span class="num">' + DADOS.macros.length + '</span><span class="lbl">Macroprocessos</span><span class="sub">' + DADOS.subs.length + ' subprocessos · ' + DADOS.ativs.length + ' atividades · ' + DADOS.tarefas.length + ' tarefas</span></div>' +
      '<div class="kpi" title="Total de processos identificados no portfólio de processos da Companhia."><span class="num">' + procs.length + '</span><span class="lbl">Processos identificados</span><span class="sub">no portfólio atual</span></div>' +
      '<div class="kpi ok" title="Processos com todos os marcos do mapeamento (M1–M10) concluídos."><span class="num">' + concl + '</span><span class="lbl">Mapeamentos concluídos</span><span class="sub">' + andamento + ' em andamento</span></div>' +
      '<div class="kpi" title="Percentual médio de execução do mapeamento entre todos os processos do portfólio."><span class="num">' + media + '%</span><span class="lbl">Avanço médio</span><span class="sub">do mapeamento do portfólio</span></div>' +
      '<div class="kpi ' + (criticos ? 'erro' : 'ok') + '" title="Riscos classificados como Alto ou Extremo, ainda não encerrados."><span class="num">' + criticos + '</span><span class="lbl">Riscos críticos abertos</span><span class="sub">nível Alto ou Extremo</span></div>' +
      '</div>' +
      '<section class="pp-sec" id="sec-cadeia"><div class="pp-sec-h"><h2>Cadeia de Valor Integrada</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="cadeia">' +
      '<aside class="cv-aside cv-missao"><h3><i class="fas fa-flag" aria-hidden="true"></i> Missão</h3><p>' + esc(INSTITUCIONAL.missao) + '</p><h3><i class="fas fa-eye" aria-hidden="true"></i> Visão</h3><p>' + esc(INSTITUCIONAL.visao) + '</p></aside>' +
      '<div class="cv-centro">' + blocoCadeia('Macroprocessos Gerenciais', 'cat-gerencial', 'fa-compass', ger) +
      blocoCadeia('Macroprocessos Finalísticos — entrega de valor à sociedade', 'cat-finalistico', 'fa-hand-holding-heart', fin) + '</div>' +
      '<aside class="cv-aside cv-proposito"><h3><i class="fas fa-hand-holding-heart" aria-hidden="true"></i> Propósito</h3><p>' + esc(INSTITUCIONAL.proposito) + '</p></aside>' +
      '<div class="cv-suporte">' + blocoCadeia('Macroprocessos de Suporte', 'cat-suporte', 'fa-gears', sup) + '</div>' +
      '<div class="cv-valores"><strong><i class="fas fa-gem" aria-hidden="true"></i> Valores</strong>' + INSTITUCIONAL.valores.map(function (v) { return '<span class="cv-chip">' + esc(v) + '</span>'; }).join('') + '</div>' +
      '</div></section>';
  }

  /* ── TELA: catálogo ───────────────────────────────────────────────── */
  var filtroCat = { macro: [], status: [], marco: '', q: '', de: '', ate: '', ordem: 'codigo' };
  /* ── SELECT (br-select) ── Anatomia: 1 Input (obrigatório), 2 Button
     terciário de abrir/fechar (obrigatório), 3 List de itens (obrigatório),
     4 Modos de seleção (obrigatório — no múltiplo: total, intermediária e
     sem seleção), 5 Ícone search (opcional). O tipo simples usa br-radio;
     o tipo múltiplo usa br-checkbox e ganha o item destacado
     "Selecionar todos". ── */
  function selectHtml(cfg) {
    var id = cfg.id, multi = !!cfg.multiplo;
    var sel = cfg.selecionados || [];
    var itens = cfg.opcoes.map(function (o, i) {
      var oid = id + '-op' + i, marcado = sel.indexOf(o.v) >= 0;
      var campo = multi
        ? '<div class="br-checkbox"><input id="' + oid + '" name="' + id + '-item" type="checkbox" value="' + esc(o.v) + '"' +
          (marcado ? ' checked' : '') + '><label for="' + oid + '">' + esc(o.r) + '</label></div>'
        : '<div class="br-radio"><input id="' + oid + '" name="' + id + '-item" type="radio" value="' + esc(o.v) + '"' +
          (marcado ? ' checked' : '') + '><label for="' + oid + '">' + esc(o.r) + '</label></div>';
      return '<div class="br-item' + (marcado ? ' selected' : '') + '" tabindex="-1" role="option" aria-selected="' + marcado + '">' + campo + '</div>';
    }).join('');
    var todos = '';
    if (multi) {
      var tudo = cfg.opcoes.length > 0 && sel.length === cfg.opcoes.length;
      todos = '<div class="br-item highlighted" data-all="data-all" tabindex="-1" role="option" aria-selected="' + tudo + '">' +
        '<div class="br-checkbox"><input id="' + id + '-all" name="' + id + '-all" type="checkbox"' + (tudo ? ' checked' : '') + '>' +
        '<label for="' + id + '-all">' + (tudo ? 'Desselecionar todos' : 'Selecionar todos') + '</label></div></div>';
    }
    return '<div class="br-select"' + (multi ? ' multiple="multiple"' : '') + ' data-select="' + esc(cfg.chave) + '">' +
      '<div class="br-input has-icon"><label for="' + id + '">' + esc(cfg.rotulo) + '</label>' +
      '<div class="input-group"><div class="input-icon"><i class="fas fa-search" aria-hidden="true"></i></div>' +
      '<input id="' + id + '" type="text" placeholder="' + esc(cfg.placeholder || '') + '" autocomplete="off"' +
      ' role="combobox" aria-expanded="false" aria-autocomplete="list" aria-controls="' + id + '-list"' +
      (multi ? ' aria-multiselectable="true"' : '') + '></div>' +
      '<button class="br-button" type="button" data-trigger="data-trigger" tabindex="-1" aria-label="Exibir lista"' +
      ' aria-controls="' + id + '-list" aria-expanded="false"><i class="fas fa-angle-down" aria-hidden="true"></i></button></div>' +
      '<div class="br-list" id="' + id + '-list" role="listbox" tabindex="-1" aria-label="Lista de opções">' +
      todos + itens + '</div></div>';
  }
  /* Ordenação do portfólio (Componente Radio): escolha mutuamente
     exclusiva, por isso caixa de opção e não caixa de seleção. O grupo usa
     fieldset + legend para o leitor de tela identificar o propósito do
     agrupamento, com uma opção marcada por padrão (checked). */
  var ORDENS = [
    { v: 'codigo', r: 'Código' },
    { v: 'nome', r: 'Nome (A–Z)' },
    { v: 'prazo', r: 'Prazo mais próximo' },
    { v: 'avanco', r: 'Maior avanço' }
  ];
  function ordenarCat(lista) {
    var c = lista.slice();
    if (filtroCat.ordem === 'nome') c.sort(function (a, b) { return String(a.Nome).localeCompare(String(b.Nome), 'pt-BR'); });
    else if (filtroCat.ordem === 'prazo') c.sort(function (a, b) {
      return String(a.Prazo_Previsto || '9999').localeCompare(String(b.Prazo_Previsto || '9999'));
    });
    else if (filtroCat.ordem === 'avanco') c.sort(function (a, b) { return pctNorm(b.Percentual) - pctNorm(a.Percentual); });
    else c.sort(function (a, b) { return String(a.Codigo).localeCompare(String(b.Codigo), 'pt-BR', { numeric: true }); });
    return c;
  }
  function filtroOrdemHtml() {
    return '<fieldset class="radio-group" id="fOrdemGroup">' +
      '<legend class="label">Ordenar por</legend>' +
      '<p class="help-text" id="fOrdemAjuda">Vale para os cartões e para a paginação abaixo.</p>' +
      '<div class="radio-inline">' +
      ORDENS.map(function (o) {
        return '<div class="br-radio small"><input id="fOrdem-' + o.v + '" type="radio" name="fOrdem" value="' + o.v + '"' +
          (filtroCat.ordem === o.v ? ' checked' : '') + ' aria-describedby="fOrdemAjuda">' +
          '<label for="fOrdem-' + o.v + '">' + esc(o.r) + '</label></div>';
      }).join('') +
      '</div></fieldset>';
  }
  var STATUS_MAPEAMENTO = ['Não iniciado', 'Em andamento', 'Concluído', 'Suspenso'];
  /* Lista de opções de status (Componente Checkbox): o filtro admite
     mais de um status ao mesmo tempo — caso de uso do checkbox, e não
     do select, que só aceitava um. Traz a anatomia completa: cabeçalho
     (rótulo + informações adicionais), lista de opções com checkbox
     "pai" (estado intermediário quando a seleção é parcial) e mensagem
     de feedback contextual. Status sem nenhum processo no portfólio
     aparecem desabilitados. */
  // Conversões entre o formato do datepicker (dd/mm/aaaa) e o ISO usado
  // nos dados da planilha
  function isoBr(txt) {
    var m = /^\s*(\d{2})\/(\d{2})\/(\d{4})/.exec(String(txt || ''));
    return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
  }
  function brIso(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
    return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
  }
  function periodoTexto() {
    if (!filtroCat.de && !filtroCat.ate) return '';
    return brIso(filtroCat.de) + (filtroCat.ate ? ' até ' + brIso(filtroCat.ate) : '');
  }
  function filtroStatusHtml(base) {
    var contagem = {};
    STATUS_MAPEAMENTO.forEach(function (s) { contagem[slug(s)] = 0; });
    base.forEach(function (p) {
      var k = slug(p.Status_Mapeamento);
      if (k in contagem) contagem[k] += 1;
    });
    var disponiveis = STATUS_MAPEAMENTO.filter(function (s) { return contagem[slug(s)] > 0; });
    var marcados = filtroCat.status.length;
    var todos = marcados > 0 && marcados === disponiveis.length;
    return '<div class="checkbox-group" id="fStatusGroup">' +
      '<p class="label" id="fStatusLabel">Status do mapeamento</p>' +
      '<p class="text-down-01">Selecione um ou mais status. Sem seleção, o portfólio mostra todos.</p>' +
      '<ul class="checkbox-list horizontal" role="group" aria-labelledby="fStatusLabel">' +
      '<li><div class="br-checkbox"><input id="fStatus-todos" name="fStatus-todos" type="checkbox" data-parent="status"' +
      (todos ? ' checked' : '') + (marcados && !todos ? ' indeterminate' : '') + '>' +
      '<label for="fStatus-todos">Todos os status</label></div></li>' +
      STATUS_MAPEAMENTO.map(function (s) {
        var v = slug(s), n = contagem[v], off = n === 0;
        return '<li><div class="br-checkbox' + (off ? ' disabled' : '') + '">' +
          '<input id="fStatus-' + v + '" name="fStatus-' + v + '" type="checkbox" value="' + v + '" data-child="status"' +
          (filtroCat.status.indexOf(v) >= 0 ? ' checked' : '') + (off ? ' disabled' : '') + '>' +
          '<label for="fStatus-' + v + '">' + esc(s) + ' (' + n + ')</label></div></li>';
      }).join('') + '</ul></div>';
  }
  function ligarFiltroStatus(el, aoMudar) {
    var grupo = el.querySelector('#fStatusGroup');
    if (!grupo) return;
    var pai = grupo.querySelector('input[data-parent]');
    var filhos = Array.prototype.slice.call(grupo.querySelectorAll('input[data-child]'));
    var ativos = filhos.filter(function (c) { return !c.disabled; });
    // estado intermediário: alguns filhos marcados, mas não todos
    var marcados = ativos.filter(function (c) { return c.checked; }).length;
    pai.indeterminate = marcados > 0 && marcados < ativos.length;
    pai.onchange = function () {
      filtroCat.status = pai.checked ? ativos.map(function (c) { return c.value; }) : [];
      aoMudar();
    };
    filhos.forEach(function (c) {
      c.onchange = function () {
        filtroCat.status = ativos.filter(function (x) { return x.checked; }).map(function (x) { return x.value; });
        aoMudar();
      };
    });
  }
  function cardProcesso(p) {
    return '<a class="proc-card" href="#/p/' + encodeURIComponent(p.Codigo) + '">' +
      '<div class="topo"><div><span class="cod" style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(codDisp(p.Codigo)) + '</span>' +
      '<div class="nome">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
      '<div class="pp-muted" style="font-size:var(--fs-sm);margin-top:4px">' +
      [p.Area_Responsavel || '', marcoRotulo(marcoAtual(p))].filter(Boolean).map(esc).join(' · ') + '</div>' +
      '<div class="rodape">' + barraPct(p.Percentual) + '</div></a>';
  }
  function catFiltrada() {
    return DADOS.procs.filter(function (p) {
      if (filtroCat.macro.length && filtroCat.macro.indexOf(p.Macroprocesso) < 0) return false;
      if (filtroCat.status.length && filtroCat.status.indexOf(slug(p.Status_Mapeamento)) < 0) return false;
      // filtro por marco: mostra os processos que estão NAQUELE marco (o
      // marco mais avançado que já concluíram), de M1 a M10.
      if (filtroCat.marco && String(marcoAtual(p)) !== filtroCat.marco) return false;
      if (filtroCat.de || filtroCat.ate) {
        var pz = p.Prazo_Previsto || '';
        if (!pz) return false;
        if (filtroCat.de && pz < filtroCat.de) return false;
        if (filtroCat.ate && pz > filtroCat.ate) return false;
      }
      if (filtroCat.q) {
        var q = filtroCat.q.toLowerCase();
        if ((p.Codigo + ' ' + p.Nome + ' ' + (p.Descricao || '')).toLowerCase().indexOf(q) < 0) return false;
      }
      return true;
    });
  }
  // Resultados (feedback + grade + paginação) redesenhados isoladamente,
  // para que a lista aberta de um select múltiplo continue aberta enquanto
  // o usuário marca várias opções em tempo real.
  function catResultadosHtml(lista) {
    return (!lista.length && filtroCat.status.length ? '<span class="feedback warning" role="alert"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i>Nenhum processo com os status selecionados.</span>' : '') +
      (lista.length ? '<div class="proc-grid">' + pagFatia('catalogo', ordenarCat(lista), 6).map(cardProcesso).join('') + '</div>' +
        paginacaoHtml('catalogo', lista.length, 'processos', [6, 12, 24, 48])
        : '<p class="pp-vazio">Nenhum processo corresponde aos filtros. Limpe os filtros para ver todos.</p>');
  }
  function renderResultados() {
    var box = $('#catResultados');
    if (!box) return renderCatalogo();
    box.innerHTML = catResultadosHtml(catFiltrada());
    ligarPaginacao(box, renderResultados);
    // O rodapé do painel vive fora de #catResultados; a visibilidade do
    // botão "Limpar filtros" acompanha cada mudança de filtro.
    var rodape = $('#fRodapeFiltros');
    if (rodape) rodape.hidden = !catTemFiltro();
  }
  // Algum filtro ativo? Controla a exibição do botão "Limpar filtros".
  function catTemFiltro() {
    return !!(filtroCat.macro.length || filtroCat.status.length || filtroCat.marco ||
      filtroCat.q || filtroCat.de || filtroCat.ate);
  }
  function renderCatalogo() {
    var el = $('#viewCatalogo');
    var lista = catFiltrada();
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Portfólio de processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      /* Painel único de filtros: uma linha de campos (Select, Select,
         DateTimePicker, Input de busca) e, após um Divider, as opções de
         recorte (Checkbox de status) e de ordenação (Radio). */
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros do portfólio">' +
      '<div class="filtros-campos">' +
      selectHtml({ chave: 'macro', id: 'fMacro', rotulo: 'Macroprocesso', multiplo: true,
        placeholder: 'Todos os macroprocessos', selecionados: filtroCat.macro,
        opcoes: DADOS.macros.map(function (m) { return { v: m.Codigo, r: (m._cod || m.Codigo) + ' — ' + m.Nome }; }) }) +
      selectHtml({ chave: 'marco', id: 'fMarco', rotulo: 'Marco do mapeamento',
        placeholder: 'Todos os marcos', selecionados: filtroCat.marco ? [filtroCat.marco] : [],
        opcoes: MARCOS_ROTULOS.map(function (r, i) { return { v: String(i + 1), r: 'M' + (i + 1) + ' - ' + r }; }) }) +
      '<div class="br-datetimepicker" data-mode="range" data-type="text">' +
      '<div class="br-input has-icon">' +
      '<label for="fPrazo">Prazo previsto entre</label>' +
      '<input id="fPrazo" type="text" placeholder="dd/mm/aaaa até dd/mm/aaaa" value="' + esc(periodoTexto()) + '" data-input>' +
      '<button class="br-button circle small" type="button" id="fPrazoBtn" aria-label="Abrir calendário"><i class="fas fa-calendar-alt" aria-hidden="true"></i></button>' +
      '</div></div>' +
      '<div class="br-input has-icon">' +
      '<label for="fBusca">Buscar no portfólio</label>' +
      '<div class="input-group"><div class="input-icon"><i class="fas fa-search" aria-hidden="true"></i></div>' +
      '<input type="search" id="fBusca" placeholder="Código ou nome do processo" value="' + esc(filtroCat.q) + '">' +
      '</div></div>' +
      '</div>' +
      '<span class="br-divider" role="presentation"></span>' +
      '<div class="filtros-opcoes">' + filtroStatusHtml(DADOS.procs) + filtroOrdemHtml() + '</div>' +
      '<div class="filtros-rodape" id="fRodapeFiltros"' +
      (catTemFiltro() ? '' : ' hidden') + '><button class="br-button secondary small" type="button" id="fLimparTudo"><i class="fas fa-rotate-left" aria-hidden="true"></i> Limpar filtros</button></div>' +
      '</section>' +
      '<div id="catResultados">' + catResultadosHtml(lista) + '</div>';
    ligarPaginacao(el, renderResultados);
    ligarFiltroStatus(el, renderResultados);
    if (window.PPDateTimePicker) window.PPDateTimePicker.init(el);
    var prazo = $('#fPrazo');
    if (prazo) prazo.onchange = function () {
      var partes = String(this.value || '').split(' até ');
      filtroCat.de = isoBr(partes[0]);
      filtroCat.ate = isoBr(partes[1]);
      renderCatalogo();
    };
    var limpar = $('#fLimparTudo');
    if (limpar) limpar.onclick = function () {
      filtroCat.macro = []; filtroCat.status = []; filtroCat.marco = '';
      filtroCat.q = ''; filtroCat.de = ''; filtroCat.ate = '';
      PAG.catalogo.pag = 1;
      renderCatalogo();
    };
    // Instancia o comportamento do Select nos dois filtros; o retorno de
    // chamada recebe os valores selecionados (evento onChange).
    window.BRSelectInit(el, function (chave, valores) {
      if (chave === 'macro') filtroCat.macro = valores;
      else filtroCat.marco = valores[0] || '';
      PAG.catalogo.pag = 1;
      renderResultados();
    });
    $all('#fOrdemGroup input[type="radio"]').forEach(function (r) {
      r.onchange = function () { filtroCat.ordem = this.value; PAG.catalogo.pag = 1; renderResultados(); };
    });
    $('#fBusca').oninput = function () { filtroCat.q = this.value; PAG.catalogo.pag = 1; renderResultados(); };
  }

  /* ── TELA: detalhe (mp | p | sp | a) ──────────────────────────────── */
  function secVinculos(nivel, codigo, semIndicadores) {
    var docs = vinculados('docs', nivel, codigo);
    var riscos = vinculados('riscos', nivel, codigo);
    var inds = semIndicadores ? [] : vinculados('inds', nivel, codigo);
    return (semIndicadores ? '' : '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> Indicadores de desempenho</h3>' + tabelaIndsHtml(inds, false) + '</div>') +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Riscos (matriz 5×5 · P×I)</h3>' + tabelaRiscosHtml(riscos, false) + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-folder-open" aria-hidden="true"></i> Normativos e documentos vinculados</h3>' + listaDocsHtml(docs) + '</div>';
  }
  function renderDetalhe(tipo, cod) {
    var el = $('#viewDetalhe');
    if (tipo === 'mp') {
      var m = IDX.mp[cod];
      if (!m) { el.innerHTML = naoEncontrado('Macroprocesso', cod); return; }
      var filhos = IDX.procsPorMacro[cod] || [];
      var media = filhos.length ? Math.round(filhos.reduce(function (s, p) { return s + p.Percentual; }, 0) / filhos.length) : 0;
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }, { rotulo: (m._cod || m.Codigo) + ' — ' + m.Nome }]) +
        '<section class="ficha-hero nv-mp cat-' + esc(m._cat || '') + '">' +
        '<span class="eyebrow">Macroprocesso ' + esc(m.Categoria) + '</span><h2>' + esc(m._cod || m.Codigo) + ' — ' + esc(m.Nome) + '</h2>' +
        '<div class="meta"><span>' + filhos.length + ' processos vinculados</span><span>· gerenciamento atual em ' + media + '%</span></div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do macroprocesso</h3><dl class="ficha-dl">' +
        campo('Objetivo', m.Objetivo && esc(m.Objetivo), false, 'desc') + campo('Descrição', m.Descricao && esc(m.Descricao), false, 'desc') +
        campo('Unidade Orgânica responsável', m.Unidade_Responsavel && esc(m.Unidade_Responsavel), false, 'quem') +
        campo('Unidades orgânicas corresponsáveis', chips(m.Unidades_Corresponsaveis), false, 'quem') +
        campo('Entregas (produtos/serviços)', chips(m.Entregas), true, 'valor') +
        campo('Beneficiários', chips(m.Clientes_Beneficiarios), false, 'valor') +
        campo('Partes interessadas', chips(m.Partes_Interessadas), false, 'valor') +
        campo('Sistemas utilizados', chips(m.Sistemas, 'fa-desktop'), false, 'tecnico') +
        (m.Observacoes ? campo('Observações', esc(m.Observacoes), true) : '') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(m.Imagem_Bizagi, m.Nome) + '</div>' +
        secVinculos('Macroprocesso', cod) +
        '</div><aside>' +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Processos vinculados</h3>' +
        (filhos.length ? filhos.map(function (p) {
          return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="#/p/' + encodeURIComponent(p.Codigo) + '">' +
            '<div class="topo"><div><span style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(codDisp(p.Codigo)) + '</span>' +
            '<div class="nome" style="font-size:var(--fs-sm)">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
            '<div class="rodape">' + barraPct(p.Percentual) + '</div></a>';
        }).join('') : '<p class="pp-vazio">Nenhum processo cadastrado.</p>') + '</div></aside></div>';
      return;
    }
    if (tipo === 'p') {
      var p = IDX.p[cod];
      if (!p) { el.innerHTML = naoEncontrado('Processo', cod); return; }
      var mp = IDX.mp[p.Macroprocesso];
      var subs = IDX.subsPorPai[cod] || [];
      // Um processo pode não ter subprocesso e ainda assim ter atividades (e
      // tarefas) ligadas direto a ele — é o que ativsDiretas cobre.
      var ativsDiretas = atividadesDe(cod);
      var totalAtivs = contarAtividadesRecursivo(cod);
      var totalTarefas = contarTarefasRecursivo(cod);
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
          .concat(mp ? [{ rotulo: mp._cod || mp.Codigo, href: '#/mp/' + encodeURIComponent(mp.Codigo) }] : [])
          .concat([{ rotulo: codDisp(p.Codigo) + ' — ' + p.Nome }])) +
        '<section class="ficha-hero nv-p">' +
        '<span class="eyebrow">Processo' + (mp ? ' · ' + esc(mp.Categoria) : '') + '</span>' +
        '<h2>' + esc(codDisp(p.Codigo)) + ' — ' + esc(p.Nome) + '</h2>' +
        '<div class="meta">' + tagStatus(p.Status_Mapeamento) +
        '<span>Mapeamento em <strong>' + p.Percentual + '%</strong></span>' +
        '<span>· ' + plural(contarSubprocessosRecursivo(cod), 'subprocesso', 'subprocessos') + ' · ' +
        plural(totalAtivs, 'atividade', 'atividades') + ' · ' + plural(totalTarefas, 'tarefa', 'tarefas') + '</span>' +
        (p.Area_Responsavel ? '<span>' + esc(p.Area_Responsavel) + '</span>' : '') +
        (p.Processo_SEI ? '<span><i class="fas fa-file-lines" aria-hidden="true"></i> SEI ' + esc(p.Processo_SEI) + '</span>' : '') +
        '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do processo</h3><dl class="ficha-dl">' +
        campo('Objetivo', p.Objetivo && esc(p.Objetivo), false, 'desc') +
        campo('Descrição', p.Descricao && esc(p.Descricao), false, 'desc') +
        campo('Unidade Orgânica responsável', p.Area_Responsavel && esc(p.Area_Responsavel), false, 'quem') +
        campo('Unidades orgânicas corresponsáveis', chips(p.Unidades_Corresponsaveis), false, 'quem') +
        campo('Responsável no NUGEP', p.Interlocutor && esc(p.Interlocutor), false, 'quem') +
        campo('Prioridade', esc(p.Prioridade || '—'), false, 'quem') +
        campo('Complexidade', esc(p.Complexidade || '—'), false, 'quem') +
        campo('Sistemas utilizados', chips(p.Sistemas, 'fa-desktop'), false, 'tecnico') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-right-left" aria-hidden="true"></i> SIPOC</h3><div class="sipoc">' +
        '<div class="col"><h4>Fornecedores</h4><ul>' + (listar(p.Fornecedores).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Entradas</h4><ul>' + (listar(p.Entradas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col centro"><h4>Processo</h4><div style="font-weight:600">' + esc(p.Nome) + '</div></div>' +
        '<div class="col"><h4>Saídas</h4><ul>' + (listar(p.Saidas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Beneficiários</h4><ul>' + (listar(p.Beneficiarios || p.Clientes).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '</div></div>' +
        '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M10)</h3>' + marcosHtml(p) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(p.Imagem_Bizagi, p.Nome) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Atividades ligadas direto ao processo</h3>' +
        (subs.length && !ativsDiretas.length
          ? '<p class="pp-vazio">Nenhuma atividade ligada diretamente ao processo — neste caso as atividades ficam dentro dos subprocessos listados ao lado.</p>'
          : tabelaAtividadesHtml(ativsDiretas, 'Nenhuma atividade cadastrada. Quando o processo não tem subprocessos, ligue as atividades direto a ele: coluna Vinculo_Pai da aba Atividades = ' + esc(cod) + '.')) + '</div>' +
        secVinculos('Processo', cod) +
        '</div><aside>' +
        cardHierarquia(mp ? [{ tipo: 'mp', cat: mp._cat, codigo: mp._cod || mp.Codigo, nome: mp.Nome, href: '#/mp/' + encodeURIComponent(mp.Codigo) }] : []) +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos vinculados</h3>' +
        (subs.length ? subs.map(function (s) {
          return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="#/sp/' + encodeURIComponent(s.Codigo) + '"><div class="topo"><div><span class="cod">' + esc(codDisp(s.Codigo)) + '</span>' +
            '<div class="nome" style="font-size:var(--fs-sm)">' + esc(s.Nome) + '</div></div></div></a>';
        }).join('') : '<p class="pp-vazio">' + (ativsDiretas.length
          ? 'Nenhum subprocesso cadastrado — este processo se decompõe direto em atividades.'
          : 'Nenhum subprocesso cadastrado.') + '</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-forward" aria-hidden="true"></i> Próxima ação</h3>' +
        (p.Proxima_Acao ? '<p style="font-size:var(--fs-sm)">' + esc(p.Proxima_Acao) + '</p>' : '<p class="pp-vazio">—</p>') +
        (p.Pendencia ? '<div class="pp-aviso" style="margin:var(--sp2) 0 0"><strong>Pendência:</strong> ' + esc(p.Pendencia) + '</div>' : '') + '</div>' +
        '</aside></div>';
      ligarLinhasTabela();
      return;
    }
    if (tipo === 'sp') {
      var s = IDX.sp[cod];
      if (!s) { el.innerHTML = naoEncontrado('Subprocesso', cod); return; }
      var cadeiaSp = cadeiaSubprocessos(cod).slice(1); // pais SP, do mais raso ao mais fundo (exclui o próprio "s")
      cadeiaSp.reverse();
      var pp = processoDoSubprocesso(cod); var mpp = pp && IDX.mp[pp.Macroprocesso];
      var paiDireto = null, paiEhSub = String(s.Vinculo_Pai || '').indexOf('SP-') === 0;
      if (paiEhSub) paiDireto = IDX.sp[s.Vinculo_Pai]; else paiDireto = pp;
      var subsFilhos = IDX.subsPorPai[cod] || [];
      var ativs = atividadesDe(cod);
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
          .concat(mpp ? [{ rotulo: codDisp(mpp.Codigo), href: '#/mp/' + encodeURIComponent(mpp.Codigo) }] : [])
          .concat(pp ? [{ rotulo: codDisp(pp.Codigo), href: '#/p/' + encodeURIComponent(pp.Codigo) }] : [])
          .concat(cadeiaSp.map(function (sp2) { return { rotulo: sp2.Codigo, href: '#/sp/' + encodeURIComponent(sp2.Codigo) }; }))
          .concat([{ rotulo: codDisp(s.Codigo) + ' — ' + s.Nome }])) +
        '<section class="ficha-hero nv-sp">' +
        '<span class="eyebrow">Subprocesso' + (paiEhSub ? ' de ' + esc(paiDireto.Nome) + ' (subprocesso)' : pp ? ' de ' + esc(pp.Nome) : '') + '</span>' +
        '<h2>' + esc(codDisp(s.Codigo)) + ' — ' + esc(s.Nome) + '</h2>' +
        '<div class="meta"><span>' + ativs.length + ' atividades mapeadas</span>' +
        (s.Unidade_Responsavel ? '<span>· ' + esc(s.Unidade_Responsavel) + '</span>' : '') + '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do subprocesso</h3><dl class="ficha-dl">' +
        campo('Descrição', s.Descricao && esc(s.Descricao), true, 'desc') +
        campo('Objetivo', s.Objetivo && esc(s.Objetivo), true, 'desc') +
        campo('Unidade Orgânica responsável', s.Unidade_Responsavel && esc(s.Unidade_Responsavel), false, 'quem') +
        campo('Unidades orgânicas corresponsáveis', chips(s.Unidades_Corresponsaveis), false, 'quem') +
        campo('Entradas (insumos)', chips(s.Entradas, 'fa-arrow-right-to-bracket'), false, 'valor') +
        campo('Saídas (produtos)', chips(s.Saidas, 'fa-arrow-right-from-bracket'), false, 'valor') +
        campo('Sistemas', chips(s.Sistemas, 'fa-desktop'), false, 'tecnico') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos deste subprocesso</h3>' +
        (subsFilhos.length ?
          '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Subprocesso</th><th>Saídas</th><th></th></tr></thead><tbody>' +
          subsFilhos.map(function (sf) {
            return '<tr data-link><td class="cod">' + esc(codDisp(sf.Codigo)) + '</td><td><a href="#/sp/' + encodeURIComponent(sf.Codigo) + '"><strong>' + esc(sf.Nome) + '</strong></a>' +
              (sf.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(sf.Descricao) + '</div>' : '') + '</td>' +
              '<td style="font-size:var(--fs-sm)">' + (listar(sf.Saidas).map(esc).join('; ') || '—') + '</td>' +
              '<td><a class="br-button secondary small" href="#/sp/' + encodeURIComponent(sf.Codigo) + '">Abrir ficha</a></td></tr>';
          }).join('') + '</tbody></table></div>'
          : '<p class="pp-vazio">Nenhum subprocesso cadastrado dentro deste subprocesso.</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama (Bizagi · BPMN)</h3>' + diagramaHtml(s.Imagem_Bizagi, s.Nome) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Atividades (com entradas e saídas)</h3>' +
        tabelaAtividadesHtml(ativs) + '</div>' +
        secVinculos('Subprocesso', cod) +
        '</div><aside>' +
        cardHierarquia(
          (mpp ? [{ tipo: 'mp', cat: mpp._cat, codigo: mpp.Codigo, nome: mpp.Nome, href: '#/mp/' + encodeURIComponent(mpp.Codigo) }] : [])
          .concat(pp ? [{ tipo: 'p', codigo: pp.Codigo, nome: pp.Nome, href: '#/p/' + encodeURIComponent(pp.Codigo) }] : [])
          .concat(cadeiaSp.map(function (sp2) { return { tipo: 'sp', codigo: sp2.Codigo, nome: sp2.Nome, href: '#/sp/' + encodeURIComponent(sp2.Codigo) }; }))
        ) +
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
    // Pai da atividade: um subprocesso (SP-...) ou o processo (P-...) direto,
    // quando o processo não tem subprocessos. O breadcrumb e o card "Navegar
    // para" simplesmente não mostram o nível que não existe.
    var anc2 = ancestraisDaAtividade(a);
    var sp2 = anc2.sp, cadeiaSp2 = anc2.cadeiaSp, p2 = anc2.p, mp2 = anc2.mp;
    var tf3 = IDX.tarefasPorAtiv[cod] || [];
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp2 ? [{ rotulo: mp2._cod || mp2.Codigo, href: '#/mp/' + encodeURIComponent(mp2.Codigo) }] : [])
        .concat(p2 ? [{ rotulo: codDisp(p2.Codigo), href: '#/p/' + encodeURIComponent(p2.Codigo) }] : [])
        .concat(cadeiaSp2.map(function (spx) { return { rotulo: spx.Codigo, href: '#/sp/' + encodeURIComponent(spx.Codigo) }; }))
        .concat([{ rotulo: codDisp(a.Codigo) }])) +
      '<section class="ficha-hero nv-a">' +
      '<span class="eyebrow">Atividade</span>' +
      '<h2>' + esc(codDisp(a.Codigo)) + ' — ' + esc(a.Nome) + '</h2>' +
      '<div class="meta">' +
      (a.Prazo_Padrao ? '<span>· Prazo padrão: ' + esc(a.Prazo_Padrao) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da atividade</h3><dl class="ficha-dl">' +
      campo('Descrição', a.Descricao && esc(a.Descricao), true, 'desc') +
      campo('Objetivo', a.Objetivo && esc(a.Objetivo), true, 'desc') +
      campo('Executor', a.Executor && esc(a.Executor), false, 'quem') +
      campo('Entradas (insumos)', chips(a.Entradas, 'fa-arrow-right-to-bracket'), false, 'valor') +
      campo('Saídas (produtos)', chips(a.Saidas, 'fa-arrow-right-from-bracket'), false, 'valor') +
      campo('Sistemas', chips(a.Sistemas, 'fa-desktop'), false, 'tecnico') + '</dl></div>' +
      '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Tarefas</h3>' +
      (tf3.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Tarefa</th><th>Tipo</th><th>Duração</th></tr></thead><tbody>' +
        tf3.map(function (t) {
          return '<tr data-link><td class="cod">' + esc(codDisp(t.Codigo)) + '</td><td><a href="#/t/' + encodeURIComponent(t.Codigo) + '"><strong>' + esc(t.Nome) + '</strong></a>' +
            (t.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(t.Descricao) + '</div>' : '') + '</td>' +
            '<td style="font-size:var(--fs-sm)">' + esc(t.Tipo_Tarefa || '—') + '</td>' +
            '<td style="font-size:var(--fs-sm);white-space:nowrap">' + esc(t.Duracao_Estimada || '—') + '</td></tr>';
        }).join('') + '</tbody></table></div>' : '<p class="pp-vazio">Nenhuma tarefa cadastrada para esta atividade.</p>') + '</div>' +
      '</div><aside>' +
      cardHierarquia(
        (mp2 ? [{ tipo: 'mp', cat: mp2._cat, codigo: mp2._cod || mp2.Codigo, nome: mp2.Nome, href: '#/mp/' + encodeURIComponent(mp2.Codigo) }] : [])
        .concat(p2 ? [{ tipo: 'p', codigo: p2.Codigo, nome: p2.Nome, href: '#/p/' + encodeURIComponent(p2.Codigo) }] : [])
        .concat(cadeiaSp2.map(function (spx) { return { tipo: 'sp', codigo: spx.Codigo, nome: spx.Nome, href: '#/sp/' + encodeURIComponent(spx.Codigo) }; }))
      ) +
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
    var a3 = IDX.a[t.Atividade];
    var anc3 = ancestraisDaAtividade(a3);
    var sp3 = anc3.sp, cadeiaSp3 = anc3.cadeiaSp, p3 = anc3.p, mp3 = anc3.mp;
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp3 ? [{ rotulo: codDisp(mp3.Codigo), href: '#/mp/' + encodeURIComponent(mp3.Codigo) }] : [])
        .concat(p3 ? [{ rotulo: codDisp(p3.Codigo), href: '#/p/' + encodeURIComponent(p3.Codigo) }] : [])
        .concat(cadeiaSp3.map(function (spx) { return { rotulo: spx.Codigo, href: '#/sp/' + encodeURIComponent(spx.Codigo) }; }))
        .concat(a3 ? [{ rotulo: codDisp(a3.Codigo), href: '#/a/' + encodeURIComponent(a3.Codigo) }] : [])
        .concat([{ rotulo: codDisp(t.Codigo) }])) +
      '<section class="ficha-hero nv-t">' +
      '<span class="eyebrow">Tarefa</span>' +
      '<h2>' + esc(codDisp(t.Codigo)) + ' — ' + esc(t.Nome) + '</h2>' +
      '<div class="meta">' + (t.Tipo_Tarefa ? '<span><i class="fas fa-gear" aria-hidden="true"></i> ' + esc(t.Tipo_Tarefa) + '</span>' : '') +
      (t.Duracao_Estimada ? '<span>· Duração estimada: ' + esc(t.Duracao_Estimada) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da tarefa</h3><dl class="ficha-dl">' +
      campo('Descrição', t.Descricao && esc(t.Descricao), true, 'desc') +
      campo('Objetivo', t.Objetivo && esc(t.Objetivo), true, 'desc') +
      campo('Tipo (CBOK 4.0)', t.Tipo_Tarefa && esc(t.Tipo_Tarefa), false, 'tecnico') +
      campo('Sistema', t.Sistema ? chips(t.Sistema, 'fa-desktop') : null, false, 'tecnico') +
      campo('Observações', t.Observacoes && esc(t.Observacoes), true) + '</dl></div>' +
      '</div><aside>' +
      cardHierarquia(
        (mp3 ? [{ tipo: 'mp', cat: mp3._cat, codigo: mp3.Codigo, nome: mp3.Nome, href: '#/mp/' + encodeURIComponent(mp3.Codigo) }] : [])
        .concat(p3 ? [{ tipo: 'p', codigo: p3.Codigo, nome: p3.Nome, href: '#/p/' + encodeURIComponent(p3.Codigo) }] : [])
        .concat(cadeiaSp3.map(function (spx) { return { tipo: 'sp', codigo: spx.Codigo, nome: spx.Nome, href: '#/sp/' + encodeURIComponent(spx.Codigo) }; }))
        .concat(a3 ? [{ tipo: 'a', codigo: a3.Codigo, nome: a3.Nome, href: '#/a/' + encodeURIComponent(a3.Codigo) }] : [])
      ) +
      '</aside></div>';
  }
  function naoEncontrado(tipo, cod) {
    return breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: tipo + ' não encontrado' }]) +
      '<div class="pp-card"><h3>' + esc(tipo) + ' não encontrado</h3><p style="font-size:var(--fs-sm)">O código <strong>' +
      esc(codDisp(cod)) + '</strong> não existe na base atual. Verifique a planilha ou volte à lista de <a href="#/catalogo">processos</a>.</p></div>';
  }

  /* ── TELAS: documentos · riscos · indicadores · diário ────────────── */
  var filtroDoc = { tipo: '', q: '', busca: '', ordem: '', dir: '', dens: 'medium', sel: {} };
  // Campo de busca no padrão Input do DS: rótulo visível, ícone
  // ilustrativo search à esquerda e input dentro do input-group.
  function buscaCampoHtml(id, rotulo, placeholder, valor) {
    return '<div class="br-input has-icon">' +
      '<label for="' + id + '">' + esc(rotulo) + '</label>' +
      '<div class="input-group"><div class="input-icon"><i class="fas fa-search" aria-hidden="true"></i></div>' +
      '<input type="search" id="' + id + '" placeholder="' + esc(placeholder) + '" value="' + esc(valor) + '"></div></div>';
  }
  // Colunas ordenáveis da tabela de documentos (Comportamento 9:
   // Ordenação — um parâmetro por vez, sem ordenação → crescente →
   // decrescente).
  var COLS_DOC = [
    { k: 'ID', r: 'ID' }, { k: 'Titulo', r: 'Documento' },
    { k: 'Vinculo_Codigo', r: 'Vinculado a' }, { k: 'Data', r: 'Data' },
    { k: 'Situacao', r: 'Situação' }
  ];
  function renderDocumentos() {
    var el = $('#viewDocumentos');
    var tipos = {};
    DADOS.docs.forEach(function (x) { if (x.Tipo_Documento) tipos[x.Tipo_Documento] = 1; });
    var bq = filtroDoc.busca.toLowerCase();
    var lista = DADOS.docs.filter(function (x) {
      if (filtroDoc.tipo && x.Tipo_Documento !== filtroDoc.tipo) return false;
      if (filtroDoc.q && (x.ID + ' ' + x.Titulo).toLowerCase().indexOf(filtroDoc.q.toLowerCase()) < 0) return false;
      // Busca da barra de título: percorre todas as colunas exibidas.
      if (bq && COLS_DOC.map(function (c) { return x[c.k] || ''; }).join(' ').toLowerCase().indexOf(bq) < 0) return false;
      return true;
    });
    if (filtroDoc.ordem) {
      var sinal = filtroDoc.dir === 'desc' ? -1 : 1;
      lista = lista.slice().sort(function (a, b) {
        return sinal * String(a[filtroDoc.ordem] || '').localeCompare(String(b[filtroDoc.ordem] || ''), 'pt-BR', { numeric: true });
      });
    }
    var pagina = pagFatia('docs', lista, 5);
    var nSel = pagina.filter(function (x) { return filtroDoc.sel[x.ID]; }).length;
    var todosSel = pagina.length > 0 && nSel === pagina.length;
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Repositório de documentos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros do repositório de documentos">' +
      '<div class="filtros-campos">' +
      selectHtml({ chave: 'tipoDoc', id: 'fTipoDoc', rotulo: 'Tipo de documento',
        placeholder: 'Todos os tipos', selecionados: filtroDoc.tipo ? [filtroDoc.tipo] : [],
        opcoes: Object.keys(tipos).sort().map(function (t) { return { v: t, r: t }; }) }) +
      buscaCampoHtml('fBuscaDoc', 'Buscar documento', 'Título do documento', filtroDoc.q) +
      '</div></section>' +
      '<div class="br-table ' + esc(filtroDoc.dens) + '" id="tabelaDocs" data-search="data-search" data-selection="data-selection" data-collapse="data-collapse">' +
      /* Barra de Título (item 1): título + ações utilitárias. Acima de 4
         ações a spec pede menu flutuante — densidade vai no ellipsis-v. */
      '<div class="table-header">' +
      '<div class="top-bar">' +
      '<div class="table-title">Documentos publicados</div>' +
      '<div class="actions-trigger text-nowrap">' +
      '<button class="br-button circle small" type="button" id="docsDensBtn" title="Ver mais opções" data-toggle="dropdown" data-target="docsDensMenu" aria-label="Definir densidade da tabela" aria-haspopup="true" aria-expanded="false"><i class="fas fa-ellipsis-v" aria-hidden="true"></i></button>' +
      '<div class="br-list dd-target" id="docsDensMenu" role="menu" aria-labelledby="docsDensBtn" hidden>' +
      [['small', 'Densidade alta'], ['medium', 'Densidade média'], ['large', 'Densidade baixa']].map(function (o, i) {
        return (i ? '<span class="br-divider" role="presentation"></span>' : '') +
          '<button class="br-item" type="button" role="menuitem" data-density="' + o[0] + '"' +
          (filtroDoc.dens === o[0] ? ' aria-current="true"' : '') + '>' + o[1] + '</button>';
      }).join('') + '</div></div>' +
      '<div class="search-trigger"><button class="br-button circle small" type="button" id="docsSearchBtn" data-toggle="search" aria-label="Abrir busca" aria-controls="docsSearchInput" aria-expanded="false"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '</div>' +
      /* Busca (Comportamento 6): cobre a barra de título enquanto ativa. */
      '<div class="search-bar' + (filtroDoc.busca ? ' show' : '') + '">' +
      '<div class="br-input"><label for="docsSearchInput">Buscar na tabela</label>' +
      '<input id="docsSearchInput" type="search" placeholder="Buscar na tabela" value="' + esc(filtroDoc.busca) + '">' +
      '<button class="br-button circle small" type="button" aria-label="Buscar"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '<button class="br-button circle small" type="button" data-dismiss="search" aria-label="Fechar busca"><i class="fas fa-times" aria-hidden="true"></i></button>' +
      '</div>' +
      /* Barra Contextual (item 2): surge sob a barra de título ao
         selecionar linhas, com contagem e ações contextuais. */
      '<div class="selected-bar' + (nSel ? ' show' : '') + '">' +
      '<div class="info"><span class="count">' + nSel + '</span><span class="text">' + (nSel === 1 ? 'item selecionado' : 'itens selecionados') + '</span></div>' +
      '<button class="br-button circle small" type="button" id="docsSelExport" aria-label="Exportar seleção em CSV" title="Exportar seleção em CSV"><i class="fas fa-file-csv" aria-hidden="true"></i></button>' +
      '<button class="br-button circle small" type="button" id="docsSelClear" aria-label="Limpar seleção" title="Limpar seleção"><i class="fas fa-times" aria-hidden="true"></i></button>' +
      '</div></div>' +
      '<div class="responsive"><table><caption>Documentos publicados</caption><thead><tr>' +
      '<td class="column-collapse" aria-hidden="true"></td>' +
      '<th class="column-checkbox" scope="col"><div class="br-checkbox hidden-label">' +
      '<input id="docsCheckAll" name="docsCheckAll" type="checkbox" aria-label="Selecionar tudo"' + (todosSel ? ' checked' : '') + '>' +
      '<label for="docsCheckAll">Selecionar todas as linhas</label></div></th>' +
      COLS_DOC.map(function (c) {
        var dir = filtroDoc.ordem === c.k ? filtroDoc.dir : '';
        var ic = dir === 'asc' ? 'fa-sort-up' : dir === 'desc' ? 'fa-sort-down' : 'fa-sort';
        return '<th scope="col"' + (dir ? ' aria-sort="' + (dir === 'asc' ? 'ascending' : 'descending') + '"' : '') + '>' +
          '<button type="button" class="sort-btn" data-sort="' + c.k + '" aria-label="Ordenar por ' + esc(c.r) + '">' +
          esc(c.r) + '<i class="fas ' + ic + '" aria-hidden="true"></i></button></th>';
      }).join('') + '</tr></thead><tbody>' +
      (pagina.length ? pagina.map(function (x, i) {
        var tit = x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '<span class="sr-only"> (abre em nova aba)</span></a>' : esc(x.Titulo);
        var cid = 'doc-col-' + i, chk = 'doc-chk-' + i, marcada = !!filtroDoc.sel[x.ID];
        return '<tr' + (marcada ? ' class="is-selected"' : '') + '>' +
          '<td class="column-collapse"><button class="br-button circle small" type="button" data-toggle="collapse" data-target="' + cid + '" aria-expanded="false" aria-controls="' + cid + '" aria-label="Expandir ou retrair ' + esc(x.Titulo) + '"><i class="fas fa-chevron-down" aria-hidden="true"></i></button></td>' +
          '<td><div class="br-checkbox hidden-label"><input id="' + chk + '" name="' + chk + '" type="checkbox" data-doc="' + esc(x.ID) + '" aria-label="Selecionar ' + esc(x.Titulo) + '"' + (marcada ? ' checked' : '') + '><label for="' + chk + '">Selecionar linha</label></div></td>' +
          '<td class="cod" data-th="ID">' + esc(x.ID) + '</td>' +
          '<td data-th="Documento"><strong>' + tit + '</strong><div class="pp-muted" style="font-size:var(--fs-sm)">' +
          esc(x.Tipo_Documento || '') + (x.Versao ? ' · v' + esc(x.Versao) : '') + '</div></td>' +
          '<td data-th="Vinculado a">' + linkVinculos(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' +
          '<td data-th="Data">' + fmtData(x.Data) + '</td>' +
          '<td data-th="Situação">' + esc(x.Situacao || '—') + '</td></tr>' +
          '<tr class="collapse"><td id="' + cid + '" colspan="7" hidden>' +
          '<div class="br-list" role="list">' +
          '<div class="br-item" role="listitem"><strong>Tipo:</strong> ' + esc(x.Tipo_Documento || '—') + '</div>' +
          '<div class="br-item" role="listitem"><strong>Versão:</strong> ' + esc(x.Versao || '—') + '</div>' +
          '<div class="br-item" role="listitem"><strong>Situação:</strong> ' + esc(x.Situacao || '—') + '</div>' +
          '</div></td></tr>';
      }).join('') : '<tr><td colspan="7" class="pp-vazio">Nenhum documento corresponde aos filtros.</td></tr>') +
      '</tbody></table></div>' +
      '<div class="table-footer">' + paginacaoHtml('docs', lista.length, 'documentos') + '</div></div>';
    ligarPaginacao(el, renderDocumentos);
    window.BRSelectInit(el, function (chave, valores) {
      filtroDoc.tipo = valores[0] || '';
      PAG.docs.pag = 1;
      renderDocumentos();
    });
    $('#fBuscaDoc').oninput = function () { filtroDoc.q = this.value; PAG.docs.pag = 1; renderDocumentos(); var n = $('#fBuscaDoc'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };

    /* Comportamentos do Componente Table: densidade, busca na barra de
       título, ordenação por coluna, seleção de linhas com barra contextual
       e expansão de linha. */
    var tabela = $('#tabelaDocs', el);
    $all('[data-density]', tabela).forEach(function (b) {
      b.onclick = function () { filtroDoc.dens = b.getAttribute('data-density'); renderDocumentos(); };
    });
    var cabec = tabela.querySelector('.table-header');
    var barraBusca = tabela.querySelector('.search-bar');
    var campoBusca = $('#docsSearchInput', tabela);
    function abrirBusca(v) {
      barraBusca.classList.toggle('show', v);
      cabec.classList.toggle('show', v);
      $('#docsSearchBtn', tabela).setAttribute('aria-expanded', String(v));
      if (v) campoBusca.focus();
    }
    if (filtroDoc.busca) { cabec.classList.add('show'); campoBusca.focus(); campoBusca.setSelectionRange(campoBusca.value.length, campoBusca.value.length); }
    $('#docsSearchBtn', tabela).onclick = function () { abrirBusca(true); };
    tabela.querySelector('[data-dismiss="search"]').onclick = function () {
      filtroDoc.busca = ''; PAG.docs.pag = 1; renderDocumentos();
    };
    campoBusca.oninput = function () { filtroDoc.busca = this.value; PAG.docs.pag = 1; renderDocumentos(); };
    campoBusca.onkeydown = function (ev) { if (ev.key === 'Escape') { filtroDoc.busca = ''; PAG.docs.pag = 1; renderDocumentos(); } };
    $all('.sort-btn', tabela).forEach(function (b) {
      b.onclick = function () {
        var c = b.getAttribute('data-sort');
        if (filtroDoc.ordem !== c) { filtroDoc.ordem = c; filtroDoc.dir = 'asc'; }
        else if (filtroDoc.dir === 'asc') filtroDoc.dir = 'desc';
        else { filtroDoc.ordem = ''; filtroDoc.dir = ''; }
        renderDocumentos();
      };
    });
    $all('input[data-doc]', tabela).forEach(function (c) {
      c.onchange = function () { filtroDoc.sel[c.getAttribute('data-doc')] = c.checked; renderDocumentos(); };
    });
    $('#docsCheckAll', tabela).onchange = function () {
      var on = this.checked;
      pagina.forEach(function (x) { filtroDoc.sel[x.ID] = on; });
      renderDocumentos();
    };
    var limparSel = $('#docsSelClear', tabela);
    if (limparSel) limparSel.onclick = function () { filtroDoc.sel = {}; renderDocumentos(); };
    var exportSel = $('#docsSelExport', tabela);
    if (exportSel) exportSel.onclick = function () {
      var linhas = lista.filter(function (x) { return filtroDoc.sel[x.ID]; });
      baixarCsv('documentos-selecionados.csv',
        [COLS_DOC.map(function (c) { return c.r; })].concat(linhas.map(function (x) {
          return COLS_DOC.map(function (c) { return String(x[c.k] == null ? '' : x[c.k]); });
        })));
    };
    // A expansão de linha fica a cargo do listener global de
    // data-toggle="collapse" (govbr-ui.js), que já alterna hidden,
    // aria-expanded e o ícone chevron — um handler local aqui dispararia
    // junto e reverteria o estado no mesmo clique.
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
      '<div class="pp-card"><h3><i class="fas fa-border-all" aria-hidden="true"></i> Matriz 5×5 (Impacto ↑ × Probabilidade →)</h3>' +
      '<div class="matriz" role="img" aria-label="Matriz de riscos cinco por cinco">' + celulas + '</div>' +
      '<div class="matriz-legenda">' + ['Baixo', 'Moderado', 'Alto', 'Extremo'].map(function (c) { return tagNivel(c); }).join('') +
      '<span class="pp-muted">Clique em um risco para ver os detalhes na tabela.</span></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Todos os riscos</h3>' +
      tabelaRiscosHtml(pagFatia('riscos', DADOS.riscos, 5), true) +
      paginacaoHtml('riscos', DADOS.riscos.length, 'riscos') + '</div>';
    ligarPaginacao(el, renderRiscos);
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
      '<div class="pp-card">' + tabelaIndsHtml(pagFatia('inds', DADOS.inds, 5), true) +
      paginacaoHtml('inds', DADOS.inds.length, 'indicadores') + '</div>';
    ligarPaginacao(el, renderIndicadores);
  }
  /* ── TELA: metodologia ────────────────────────────────────────────── */
  /* ── GRÁFICOS (SVG puro, sem dependências; cores do DS gov.br) ────── */
  // Paleta categórica dos gráficos: as sete matizes do sistema de camadas
  // (azul, verde, laranja, turquesa, índigo, bronze, grafite) mais o
  // vermelho de erro — todas degraus oficiais da paleta gov.br, todas com
  // 4,5:1 ou mais sobre o card branco. Matizes diferentes entre si, para
  // que séries vizinhas nunca se confundam.
  var PAL = ['#0c326f', '#168821', '#c05600', '#0081a1', '#4a50c4', '#776017', '#555555', '#b50909'];
  function svgWrap(titulo, conteudo, vb, altura, legenda) {
    return '<figure class="graf"><figcaption>' + esc(titulo) +
      (legenda ? ' <i class="fas fa-circle-info graf-info" tabindex="0" data-tooltip-text="' + esc(legenda) + '" aria-label="O que este gráfico mostra e como lê-lo"></i>' : '') + '</figcaption>' +
      '<svg viewBox="' + vb + '" role="img" aria-label="' + esc(titulo) + (legenda ? '. ' + esc(legenda) : '') + '" ' +
      (altura ? 'style="height:' + altura + 'px"' : '') + '>' + conteudo + '</svg></figure>';
  }
  function grafDonut(titulo, dados, legenda) {           // [{rotulo, valor, cor}]
    var total = dados.reduce(function (a, b) { return a + b.valor; }, 0);
    if (!total) return svgWrap(titulo, '<text x="150" y="90" text-anchor="middle" font-size="12" fill="#636363">Sem dados</text>', '0 0 300 180', null, legenda);
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
    return svgWrap(titulo, arcos + '<text x="' + cx + '" y="' + (cy + 5) + '" text-anchor="middle" font-size="20" font-weight="700" fill="#1b1b1b">' + total + '</text>' + leg, '0 0 380 180', null, legenda);
  }
  function grafBarras(titulo, dados, sufixo, meta, legenda) {   // [{rotulo, valor, cor, href}]
    if (!dados.length) return svgWrap(titulo, '', '0 0 480 120', null, legenda);
    var max = Math.max.apply(null, dados.map(function (d) { return d.valor; }).concat(meta ? [meta] : [1])) || 1;
    var lw = 132, bw = 320, h = 26, alt = dados.length * h + 24;
    var barras = dados.map(function (d, i) {
      var y = i * h + 8, w = Math.max(2, d.valor / max * bw);
      var navAttr = d.href ? ' data-nav="1" data-href="' + esc(d.href) + '"' : '';
      return '<text x="0" y="' + (y + 13) + '" font-size="11" fill="#333">' + esc(String(d.rotulo).slice(0, 22)) + '</text>' +
        '<rect x="' + lw + '" y="' + y + '" width="' + bw + '" height="15" rx="3" fill="#f0f0f0"></rect>' +
        '<rect x="' + lw + '" y="' + y + '" width="' + w.toFixed(1) + '" height="15" rx="3" fill="' + (d.cor || '#0c326f') + '"' + navAttr + '>' +
        '<title>' + esc(d.rotulo) + ': ' + d.valor + (sufixo || '') + (d.href ? ' — clique para abrir' : '') + '</title></rect>' +
        '<text x="' + (lw + bw + 6) + '" y="' + (y + 12) + '" font-size="11" font-weight="700" fill="#1b1b1b">' + d.valor + (sufixo || '') + '</text>';
    }).join('');
    var linhaMeta = meta ? '<line x1="' + (lw + meta / max * bw) + '" y1="2" x2="' + (lw + meta / max * bw) + '" y2="' + (alt - 16) +
      '" stroke="#b50909" stroke-width="1.5" stroke-dasharray="4 3"><title>Meta: ' + meta + (sufixo || '') + '</title></line>' : '';
    return svgWrap(titulo, barras + linhaMeta, '0 0 500 ' + alt, null, legenda);
  }
  function grafFunil(titulo, etapas, legenda) {          // [{rotulo, valor}]
    var max = etapas[0] ? etapas[0].valor : 1, h = 30, alt = etapas.length * h + 12;
    var corpo = etapas.map(function (e, i) {
      var w = Math.max(6, (e.valor / (max || 1)) * 300), y = i * h + 4;
      return '<rect x="' + (150 - w / 2 + 150) + '" y="' + y + '" width="' + w.toFixed(1) + '" height="22" rx="3" fill="' + PAL[i % PAL.length] + '" opacity="0.9">' +
        '<title>' + esc(e.rotulo) + ': ' + e.valor + '</title></rect>' +
        '<text x="0" y="' + (y + 16) + '" font-size="11" fill="#333">' + esc(e.rotulo) + '</text>' +
        '<text x="' + (300 + 150 + w / 2 + 8) + '" y="' + (y + 16) + '" font-size="11" font-weight="700" fill="#1b1b1b">' + e.valor + '</text>';
    }).join('');
    return svgWrap(titulo, corpo, '0 0 640 ' + alt, null, legenda);
  }
  function grafLinha(titulo, pontos, legenda) {          // [{rotulo, valor}]
    if (pontos.length < 2) return svgWrap(titulo, '<text x="200" y="70" text-anchor="middle" font-size="12" fill="#636363">Dados insuficientes</text>', '0 0 400 140', null, legenda);
    var w = 420, h = 120, pad = 30;
    var max = Math.max.apply(null, pontos.map(function (p) { return p.valor; })) || 1;
    var pts = pontos.map(function (p, i) {
      return [pad + i * ((w - pad * 2) / (pontos.length - 1)), h - 20 - (p.valor / max) * (h - 45)];
    });
    var linha = pts.map(function (p, i) { return (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1); }).join(' ');
    var area = linha + ' L' + pts[pts.length - 1][0].toFixed(1) + ' ' + (h - 20) + ' L' + pts[0][0].toFixed(1) + ' ' + (h - 20) + ' Z';
    var marcas = pts.map(function (p, i) {
      return '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="3.5" fill="#0c326f"><title>' + esc(pontos[i].rotulo) + ': ' + pontos[i].valor + '</title></circle>' +
        '<text x="' + p[0].toFixed(1) + '" y="' + (h - 6) + '" font-size="9" fill="#636363" text-anchor="middle">' + esc(pontos[i].rotulo) + '</text>';
    }).join('');
    return svgWrap(titulo, '<path d="' + area + '" fill="#0c326f" opacity="0.12"></path>' +
      '<path d="' + linha + '" fill="none" stroke="#0c326f" stroke-width="2.5"></path>' + marcas, '0 0 ' + w + ' ' + h, null, legenda);
  }
  function grafHeat(titulo, dados, legenda) {            // matriz risco compacta [{p,i,qtd}]
    var cel = 30, sz = 5 * cel, corpo = '';
    for (var i = 5; i >= 1; i--) {
      for (var pb = 1; pb <= 5; pb++) {
        var n = pb * i, cls = n >= 20 ? '#fdb8ae' : n >= 12 ? '#ffbc78' : n >= 5 ? '#fee685' : '#b7f5bd';
        var achou = dados.filter(function (d) { return d.p === pb && d.i === i; })[0];
        var x = 26 + (pb - 1) * cel, y = (5 - i) * cel;
        corpo += '<rect x="' + x + '" y="' + y + '" width="' + (cel - 3) + '" height="' + (cel - 3) + '" rx="3" fill="' + cls + '"></rect>';
        if (achou) corpo += '<text x="' + (x + cel / 2 - 1.5) + '" y="' + (y + cel / 2 + 2) + '" font-size="12" font-weight="700" fill="#1b1b1b" text-anchor="middle">' + achou.qtd + '<title>Prob. ' + pb + ' × Impacto ' + i + ': ' + achou.qtd + ' risco(s)</title></text>';
      }
      corpo += '<text x="18" y="' + ((5 - i) * cel + cel / 2) + '" font-size="10" fill="#636363" text-anchor="end">' + i + '</text>';
    }
    for (var k = 1; k <= 5; k++) corpo += '<text x="' + (26 + (k - 1) * cel + cel / 2 - 1.5) + '" y="' + (sz + 10) + '" font-size="10" fill="#636363" text-anchor="middle">' + k + '</text>';
    corpo += '<text x="88" y="' + (sz + 24) + '" font-size="10" fill="#636363">Probabilidade →</text>';
    return svgWrap(titulo, corpo, '0 0 200 ' + (sz + 30), null, legenda);
  }

  function grafGauge(titulo, valor, max, faixas, legenda) {   // faixas: [{ate, cor}] em ordem crescente
    var cx = 150, cy = 140, r = 110, r2 = 82;
    function pt(v, raio) {
      var ang = Math.PI - (Math.max(0, Math.min(max, v)) / max) * Math.PI;
      return [cx + raio * Math.cos(ang), cy - raio * Math.sin(ang)];
    }
    var faixasHtml = (faixas || []).map(function (f, i) {
      var de = i === 0 ? 0 : faixas[i - 1].ate, a = pt(de, r), b = pt(f.ate, r);
      var grande = (f.ate - de) / max > 0.5 ? 1 : 0;
      return '<path d="M' + a[0].toFixed(1) + ',' + a[1].toFixed(1) + ' A' + r + ',' + r + ' 0 ' + grande + ',1 ' + b[0].toFixed(1) + ',' + b[1].toFixed(1) + '" stroke="' + f.cor + '" stroke-width="20" fill="none" opacity=".25"></path>';
    }).join('');
    var pa = pt(0, r2), pb = pt(valor, r2);
    var grandeArco = (valor / max) > 0.5 ? 1 : 0;
    var corAtual = '#0c326f';
    (faixas || []).forEach(function (f) { if (valor <= f.ate) corAtual = corAtual === '#0c326f' ? f.cor : corAtual; });
    var ponteiro = pt(valor, r2);
    var corpo = faixasHtml +
      '<path d="M' + pa[0].toFixed(1) + ',' + pa[1].toFixed(1) + ' A' + r2 + ',' + r2 + ' 0 ' + grandeArco + ',1 ' + pb[0].toFixed(1) + ',' + pb[1].toFixed(1) + '" stroke="' + corAtual + '" stroke-width="20" fill="none" stroke-linecap="round"><title>' + esc(titulo) + ': ' + valor + '% (de ' + max + '%)</title></path>' +
      '<circle cx="' + ponteiro[0].toFixed(1) + '" cy="' + ponteiro[1].toFixed(1) + '" r="7" fill="' + corAtual + '" stroke="#fff" stroke-width="2"></circle>' +
      '<text x="' + cx + '" y="' + (cy - 18) + '" font-size="34" font-weight="700" fill="#1b1b1b" text-anchor="middle">' + valor + '%</text>' +
      '<text x="' + cx + '" y="' + (cy + 6) + '" font-size="11" fill="#636363" text-anchor="middle">de ' + max + '%</text>';
    return svgWrap(titulo, corpo, '0 0 300 160', null, legenda);
  }

  function grafBubble(titulo, pontos, rotEixoX, rotEixoY, legenda) {   // [{x,y,r,rotulo,cor,href}], x,y 0-100
    var lw = 30, bh = 210, bw = 380, mt = 10, maxY = Math.max(2, Math.max.apply(null, pontos.map(function (p) { return p.y; })) + 1);
    function px(v) { return lw + (v / 100) * bw; }
    function py(v) { return mt + bh - (v / maxY) * bh; }
    var eixos = '<line x1="' + lw + '" y1="' + mt + '" x2="' + lw + '" y2="' + (mt + bh) + '" stroke="#ccc"></line>' +
      '<line x1="' + lw + '" y1="' + (mt + bh) + '" x2="' + (lw + bw) + '" y2="' + (mt + bh) + '" stroke="#ccc"></line>' +
      [0, 25, 50, 75, 100].map(function (v) { return '<text x="' + px(v) + '" y="' + (mt + bh + 14) + '" font-size="9" fill="#636363" text-anchor="middle">' + v + '</text>'; }).join('') +
      '<text x="' + (lw + bw / 2) + '" y="' + (mt + bh + 30) + '" font-size="10" fill="#636363" text-anchor="middle">' + esc(rotEixoX) + '</text>' +
      '<text x="4" y="' + (mt + 4) + '" font-size="9" fill="#636363">' + esc(rotEixoY) + ' ↑</text>';
    var bolhas = pontos.map(function (p) {
      var navAttr = p.href ? ' data-nav="1" data-href="' + esc(p.href) + '"' : '';
      return '<circle cx="' + px(p.x).toFixed(1) + '" cy="' + py(p.y).toFixed(1) + '" r="' + (p.r || 8) + '" fill="' + (p.cor || '#0c326f') + '" fill-opacity=".78" stroke="#fff" stroke-width="1.5"' + navAttr + '>' +
        '<title>' + esc(p.rotulo) + ' — ' + esc(rotEixoX) + ': ' + p.x + '% · ' + esc(rotEixoY) + ': ' + p.y + (p.href ? ' — clique para abrir' : '') + '</title></circle>';
    }).join('');
    return svgWrap(titulo, eixos + bolhas, '0 0 ' + (lw + bw + 20) + ' ' + (mt + bh + 45), null, legenda);
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
    // marcos agregados (funil)
    var funil = MARCOS_ROTULOS.map(function (rot, i) {
      return { rotulo: 'M' + (i + 1) + ' · ' + rot, valor: procs.filter(function (p) { return simNao(valMarco(p, i)); }).length };
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

    // Mesmas cores da cadeia de valor e das fichas: azul = gerencial,
    // verde = finalístico, laranja = suporte.
    var CAT_COR = { gerencial: '#0c326f', finalistico: '#168821', suporte: '#c05600' };
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Dashboard gerencial</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="kpi-grid" style="margin-top:0">' +
      '<div class="kpi" title="Processos com todos os marcos do mapeamento (M1–M10) concluídos, em relação ao total do portfólio."><span class="num">' + cobertura + '%</span><span class="lbl">Processos publicados</span><span class="sub">' + concl + ' de ' + procs.length + ' processos</span></div>' +
      '<div class="kpi" title="Percentual médio de execução do mapeamento entre todos os processos do portfólio."><span class="num">' + media + '%</span><span class="lbl">Avanço médio</span><span class="sub">' + andam + ' em andamento</span></div>' +
      '<div class="kpi" title="Total de processos mapeados no portfólio."><span class="num">' + procs.length + '</span><span class="lbl">Processos</span><span class="sub">mapeados no portfólio</span></div>' +
      '<div class="kpi" title="Total de subprocessos mapeados no portfólio (inclusive os aninhados, subprocesso dentro de subprocesso)."><span class="num">' + DADOS.subs.length + '</span><span class="lbl">Subprocessos</span><span class="sub">mapeados no portfólio</span></div>' +
      '<div class="kpi" title="Total de atividades mapeadas no portfólio."><span class="num">' + DADOS.ativs.length + '</span><span class="lbl">Atividades</span><span class="sub">mapeadas no portfólio</span></div>' +
      '<div class="kpi" title="Total de tarefas mapeadas no portfólio (menor unidade de trabalho, CBOK 4.0)."><span class="num">' + DADOS.tarefas.length + '</span><span class="lbl">Tarefas</span><span class="sub">mapeadas no portfólio</span></div>' +
      '<div class="kpi ' + (criticos.length ? 'erro' : 'ok') + '" title="Riscos classificados como Alto ou Extremo, ainda não encerrados."><span class="num">' + criticos.length + '</span><span class="lbl">Riscos críticos</span><span class="sub">' + riscosAb.length + ' riscos abertos no total</span></div>' +
      '</div>' +
      '<div class="graf-grid">' +
      grafGauge('Avanço médio geral do portfólio', media, 100, [
        { ate: 40, cor: '#b50909' }, { ate: 70, cor: '#947100' }, { ate: 100, cor: '#168821' }],
        'Percentual médio de avanço do mapeamento entre todos os processos do portfólio. As faixas de cor indicam o estágio geral: vermelho abaixo de 40%, âmbar entre 40% e 70%, verde a partir de 70%.') +
      grafDonut('Situação do mapeamento', [
        { rotulo: 'Concluído', valor: concl, cor: '#168821' },
        { rotulo: 'Em andamento', valor: andam, cor: '#947100' },
        { rotulo: 'Não iniciado', valor: porStatus['Não iniciado'] || 0, cor: '#757575' },
        { rotulo: 'Suspenso', valor: porStatus['Suspenso'] || 0, cor: '#b50909' }],
        'Quantidade de processos em cada status de mapeamento. O tamanho de cada fatia é proporcional ao número de processos; passe o mouse sobre uma fatia para ver o total e o percentual.') +
      grafDonut('Processos por tipo (CBOK)', ['gerencial', 'finalistico', 'suporte'].map(function (c) {
        return { rotulo: c === 'finalistico' ? 'Finalístico' : c === 'gerencial' ? 'Gerencial' : 'Suporte',
          valor: procs.filter(function (p) { var m = IDX.mp[p.Macroprocesso]; return m && m._cat === c; }).length, cor: CAT_COR[c] };
      }), 'Distribuição dos processos mapeados entre os três tipos de macroprocesso do CBOK 4.0: gerencial, finalístico e de suporte.') +
      grafBarras('Avanço por macroprocesso (%) — clique para abrir', DADOS.macros.map(function (m) {
        var ps = IDX.procsPorMacro[m.Codigo] || [];
        return { rotulo: (m._cod || m.Codigo) + ' ' + m.Nome, cor: CAT_COR[m._cat] || '#0c326f', href: '#/mp/' + encodeURIComponent(m.Codigo),
          valor: ps.length ? Math.round(ps.reduce(function (s, p) { return s + p.Percentual; }, 0) / ps.length) : 0 };
      }), '%', 100, 'Percentual médio de avanço dos processos de cada macroprocesso. A linha tracejada marca a meta de 100%; clique em uma barra para abrir a ficha do macroprocesso.') +
      grafBubble('Priorização: avanço × riscos abertos por processo — clique para abrir', procs.map(function (p) {
        var rp = vinculados('riscos', 'Processo', p.Codigo).filter(function (r) { return !/encerrad/i.test(String(r.Status || '')); });
        var ativs = contarAtividadesRecursivo(p.Codigo);
        return { x: p.Percentual, y: rp.length, r: Math.max(6, Math.min(22, 6 + ativs * 1.5)),
          rotulo: codDisp(p.Codigo) + ' — ' + p.Nome, cor: p.Percentual < 40 && rp.length >= 1 ? '#b50909' : (CAT_COR[(IDX.mp[p.Macroprocesso] || {})._cat] || '#0c326f'),
          href: '#/p/' + encodeURIComponent(p.Codigo) };
      }), 'Avanço do mapeamento (%)', 'Riscos abertos', 'Cada bolha é um processo: posição horizontal = avanço do mapeamento, posição vertical = riscos abertos, tamanho = atividades mapeadas. Bolhas vermelhas (avanço baixo e algum risco aberto) merecem atenção prioritária; clique para abrir a ficha.') +
      grafFunil('Marcos concluídos no portfólio (M1 → M10)', funil, 'Quantidade de processos que já concluíram cada marco do mapeamento, do M1 ao M10. Como os marcos são sequenciais, o número tende a diminuir da primeira para a última etapa.') +
      grafLinha('Processos publicados (acumulado)', linha, 'Quantidade acumulada de processos concluídos, por mês de conclusão. Cada ponto soma os concluídos até aquele mês; passe o mouse sobre um ponto para ver o mês e o total acumulado.') +
      grafHeat('Riscos abertos por probabilidade × impacto', heatArr, 'Quantidade de riscos abertos por combinação de probabilidade e impacto (escala 1 a 5). Células mais à direita e mais acima concentram os riscos mais graves.') +
      grafBarras('Riscos abertos por categoria', (function () {
        var c = {}; riscosAb.forEach(function (r) { c[r.Categoria || '—'] = (c[r.Categoria || '—'] || 0) + 1; });
        return Object.keys(c).sort(function (a, b) { return c[b] - c[a]; }).map(function (k, i) { return { rotulo: k, valor: c[k], cor: PAL[i % PAL.length] }; });
      })(), undefined, undefined, 'Quantidade de riscos abertos, agrupados por categoria. Passe o mouse sobre uma barra para ver o total exato.') +
      '</div>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Pontos de atenção</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> Processos com prazo vencido</h3>' +
      (atrasados.length ? '<div class="br-table pp-tabela-wrap"><table class="pp-tabela"><thead><tr><th>Código</th><th>Processo</th><th>Responsável</th><th>Prazo</th><th>Avanço</th></tr></thead><tbody>' +
        atrasados.map(function (p) {
          return '<tr data-link><td class="cod">' + esc(codDisp(p.Codigo)) + '</td><td><a href="#/p/' + encodeURIComponent(p.Codigo) + '"><strong>' + esc(p.Nome) + '</strong></a></td>' +
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
    $all('#viewDashboard [data-nav]').forEach(function (el2) {
      el2.setAttribute('tabindex', '0');
      el2.setAttribute('role', 'link');
      el2.addEventListener('click', function () { location.hash = el2.getAttribute('data-href'); });
      el2.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); location.hash = el2.getAttribute('data-href'); }
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
      (met || guia ?
        '<div class="repo-oficial">' +
        (met ? '<a class="repo-oficial-card" href="' + esc(met) + '" target="_blank" rel="noopener"><i class="fas fa-scale-balanced" aria-hidden="true"></i><div><strong>Metodologia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicada na intranet/SEI</span></div><i class="fas fa-up-right-from-square seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        (guia ? '<a class="repo-oficial-card" href="' + esc(guia) + '" target="_blank" rel="noopener"><i class="fas fa-book-open" aria-hidden="true"></i><div><strong>Guia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicado na intranet/SEI</span></div><i class="fas fa-up-right-from-square seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        '</div>' : '') +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Jornada de mapeamento</h2><div class="linha" aria-hidden="true"></div></div>' +
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
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros de instrumentos, modelos e ferramentas">' +
      '<div class="filtros-campos">' +
      selectHtml({ chave: 'repoCat', id: 'repoCat', rotulo: 'Categoria',
        placeholder: 'Todas as categorias', selecionados: filtroRepo.cat ? [filtroRepo.cat] : [],
        opcoes: cats.map(function (c) { return { v: c, r: c }; }) }) +
      selectHtml({ chave: 'repoFase', id: 'repoFase', rotulo: 'Fase do ciclo',
        placeholder: 'Todas as fases', selecionados: filtroRepo.fase ? [filtroRepo.fase] : [],
        opcoes: fases.map(function (c) { return { v: c, r: c }; }) }) +
      buscaCampoHtml('repoQ', 'Buscar no repositório', 'Título, código ou descrição', filtroRepo.q) +
      '</div></section>' +
      (lista.length ? '<div class="repo-grid">' + pagFatia('repo', lista, 6).map(cardRepo).join('') + '</div>' +
        paginacaoHtml('repo', lista.length, 'itens', [6, 12, 24, 48]) : '<p class="pp-vazio">Nenhum item com esses filtros.</p>') + '</section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Metodologia em resumo</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M10)</h3>' +
      '<p style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Roteiro-padrão de cada projeto de mapeamento, do primeiro contato com a área até a publicação no repositório — passe o cursor sobre um marco para ver o que ele significa:</p>' +
      '<ul class="marcos">' + MARCOS_ROTULOS.map(function (r, i) { return '<li class="feito" title="' + esc(MARCOS_DESCRICOES[i]) + '"><span>' + esc(r) + '</span><i class="fas fa-check-circle" aria-hidden="true"></i></li>'; }).join('') + '</ul>' +
      // Faixa contínua sob os marcos: monitoramento e avaliação não são um
      // marco, são atividades permanentes que atravessam todos eles.
      '<div class="marcos-continuo">' +
      '<span class="mc-rot"><i class="fas fa-arrows-rotate" aria-hidden="true"></i> Monitoramento e avaliação</span>' +
      '<span class="mc-trilho" aria-hidden="true"></span>' +
      '<span class="mc-nota">atividades contínuas — atravessam do M1 ao M10 e seguem depois da publicação</span>' +
      '</div></div>' +
      '<div class="br-message warning" role="status"><div class="icon"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i></div><div class="content"><span class="message-title">Dados fictícios.</span> <span class="message-body">Todo o conteúdo exibido foi criado apenas para demonstrar o painel — substitua na planilha.</span></div><div class="close"><button class="br-button circle small" type="button" aria-label="Fechar a mensagem"><i class="fas fa-times" aria-hidden="true"></i></button></div></div></section>';
    ligarPaginacao(el, renderRepositorio);
    window.BRSelectInit(el, function (chave, valores) {
      if (chave === 'repoCat') filtroRepo.cat = valores[0] || '';
      else filtroRepo.fase = valores[0] || '';
      PAG.repo.pag = 1;
      renderRepositorio();
    });
    var f3 = $('#repoQ');
    if (f3) f3.oninput = function () { filtroRepo.q = this.value; PAG.repo.pag = 1; renderRepositorio(); var n = $('#repoQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
  }

  /* ── TELA: NUGEP ──────────────────────────────────────────────────── */
  function iniciais(nome) {
    var p = String(nome || '').trim().split(/\s+/);
    return (((p[0] || '')[0] || '') + ((p.length > 1 ? p[p.length - 1][0] : '') || '')).toUpperCase();
  }
  function processosDoNugep(nome) {
    var alvo = String(nome || '').trim().toLowerCase();
    if (!alvo) return [];
    return DADOS.procs.filter(function (p) {
      return String(p.Interlocutor || '').toLowerCase().indexOf(alvo) === 0;
    });
  }
  function renderNugep() {
    var el = $('#viewNugep');
    var P = DADOS.params || {};
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>NUGEP — Núcleo de Gestão Normativa e de Processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      (DADOS.nugep.length ? '<div class="nugep-grid">' + DADOS.nugep.map(function (m) {
        var meusProcs = processosDoNugep(m.Nome);
        return '<article class="nugep-card"><span class="br-avatar" title="' + esc(m.Nome) + '">' +
          '<span class="content">' + esc(iniciais(m.Nome)) + '</span></span>' +
          '<h4>' + esc(m.Nome) + '</h4>' +
          '<p class="nugep-unid"><span class="nugep-sigla">' + esc(m.Unidade_Sigla || '') + '</span></p>' +
          '<div class="nugep-contato">' +
          (m.Email ? '<a href="mailto:' + esc(m.Email) + '"><i class="fas fa-envelope" aria-hidden="true"></i> ' + esc(m.Email) + '</a>' : '') +
          (m.Telefone ? '<a href="tel:+55' + esc(String(m.Telefone).replace(/\D/g, '')) + '"><i class="fas fa-phone" aria-hidden="true"></i> ' + esc(m.Telefone) + '</a>' : '') +
          '</div>' +
          (meusProcs.length ? '<div class="nugep-procs"><b><i class="fas fa-diagram-project" aria-hidden="true"></i> Processos sob responsabilidade</b><ul>' +
            meusProcs.map(function (p) { return '<li><a href="#/p/' + encodeURIComponent(p.Codigo) + '">' + esc(codDisp(p.Codigo)) + ' — ' + esc(p.Nome) + '</a></li>'; }).join('') +
            '</ul></div>' : '') +
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
    var pagina = pagFatia('gloss', lista, 12);
    var porLetra = {};
    pagina.forEach(function (t) { var L = String(t.Termo || '').charAt(0).toUpperCase(); (porLetra[L] = porLetra[L] || []).push(t); });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Glossário de Gestão de Processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros do glossário">' +
      '<div class="filtros-campos">' +
      buscaCampoHtml('glossQ', 'Buscar termo', 'Termo, sigla ou conceito (ex.: SIPOC, KPI, raia)', filtroGloss.q) +
      selectHtml({ chave: 'glossCat', id: 'glossCat', rotulo: 'Categoria',
        placeholder: 'Todas as categorias (' + todos.length + ')', selecionados: filtroGloss.cat ? [filtroGloss.cat] : [],
        opcoes: cats.map(function (c) {
          var n = todos.filter(function (t) { return t.Categoria === c; }).length;
          return { v: c, r: c + ' (' + n + ')' };
        }) }) +
      '</div>' +
      '<span class="br-divider" role="presentation"></span>' +
      '<div class="gloss-abc" role="group" aria-label="Filtrar por letra"><button type="button" class="' + (filtroGloss.letra ? '' : 'ativo') + '" data-letra="">Todos</button>' +
      abc.map(function (L) { return '<button type="button" data-letra="' + L + '" class="' + (filtroGloss.letra === L ? 'ativo' : '') + '"' + (letrasDisp[L] ? '' : ' disabled') + '>' + L + '</button>'; }).join('') + '</div>' +
      '</section>' +
      (lista.length ? Object.keys(porLetra).sort().map(function (L) {
        return '<h3 class="gloss-letra">' + L + '</h3><div class="gloss-grid">' + porLetra[L].map(function (t) {
          return '<article class="gloss-card"><div class="gloss-topo"><h4>' + esc(t.Termo) + '</h4><span class="gloss-cat">' + esc(t.Categoria || '') + '</span></div>' +
            '<p>' + esc(t.Definicao || '') + '</p>' +
            '<div class="gloss-rodape">' + (t.Fonte ? '<span class="repo-fonte">Fonte: ' + esc(t.Fonte) + '</span>' : '') +
            (t.Termos_Relacionados ? '<span class="chip-lista">' + listar(t.Termos_Relacionados).map(function (r) { return '<button type="button" class="chip gloss-rel" data-termo="' + esc(r) + '">' + esc(r) + '</button>'; }).join('') + '</span>' : '') +
            '</div></article>';
        }).join('') + '</div>';
      }).join('') : '<p class="pp-vazio">Nenhum termo encontrado com esses filtros.</p>') +
      paginacaoHtml('gloss', lista.length, 'termos', [12, 24, 48, 96]);
    ligarPaginacao(el, renderGlossario);
    window.BRSelectInit(el, function (chave, valores) {
      filtroGloss.cat = valores[0] || '';
      PAG.gloss.pag = 1;
      renderGlossario();
    });
    var q = $('#glossQ');
    if (q) q.oninput = function () { filtroGloss.q = this.value; filtroGloss.letra = ''; PAG.gloss.pag = 1; renderGlossario(); var n = $('#glossQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
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
    var porCat = {}; pagFatia('faq', lista, 5).forEach(function (f) { (porCat[f.Categoria || 'Geral'] = porCat[f.Categoria || 'Geral'] || []).push(f); });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h2>Perguntas frequentes</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="faq-cats"><button type="button" class="chip ' + (filtroFaq ? '' : 'ativo') + '" data-cat="">Todas (' + todos.length + ')</button>' +
      cats.map(function (c) { var n = todos.filter(function (f) { return f.Categoria === c; }).length; return '<button type="button" class="chip ' + (filtroFaq === c ? 'ativo' : '') + '" data-cat="' + esc(c) + '">' + esc(c) + ' (' + n + ')</button>'; }).join('') + '</div>' +
      // Accordion do gov.br DS (br-accordion): item + header/icon/title e a
      // div .content logo depois de cada item, como manda a documentação. Sem
      // o atributo `single`, então o usuário pode abrir várias perguntas ao
      // mesmo tempo — o comportamento padrão do componente.
      Object.keys(porCat).map(function (cat, ic) {
        return '<h3 class="faq-cat-h">' + esc(cat) + '</h3>' +
          '<div class="br-accordion" id="faqAcc' + ic + '">' + porCat[cat].map(function (f, i) {
            var id = 'faq-' + ic + '-' + i;
            return '<div class="item">' +
              '<button class="header" type="button" aria-controls="' + id + '" aria-expanded="false">' +
              '<span class="icon"><i class="fas fa-angle-down" aria-hidden="true"></i></span>' +
              '<span class="title">' + esc(f.Pergunta) + '</span></button></div>' +
              '<div class="content" id="' + id + '">' + esc(f.Resposta || '') + '</div>';
          }).join('') + '</div>';
      }).join('') + paginacaoHtml('faq', lista.length, 'perguntas');
    if (window.PPUI && PPUI.iniciarAccordions) PPUI.iniciarAccordions(el);
    ligarPaginacao(el, renderFaq);
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
      mp: DADOS.macros.filter(function (m) { return bate(m.Codigo) || bate(m._cod) || bate(m.Nome) || bate(m.Descricao); }),
      p: DADOS.procs.filter(function (p) { return bate(p.Codigo) || bate(p.Nome) || bate(p.Descricao); }),
      sp: DADOS.subs.filter(function (s) { return bate(s.Codigo) || bate(s.Nome) || bate(s.Descricao); }),
      a: DADOS.ativs.filter(function (a) { return bate(a.Codigo) || bate(a.Nome) || bate(a.Descricao); }),
      t: DADOS.tarefas.filter(function (t) { return bate(t.Codigo) || bate(t.Nome) || bate(t.Descricao); }),
      doc: DADOS.docs.filter(function (x) { return bate(x.ID) || bate(x.Titulo); }),
      gl: DADOS.glossario.filter(function (t) { return bate(t.Termo) || bate(t.Definicao); }),
      rp: DADOS.repo.filter(function (i) { return bate(i.Titulo) || bate(i.Descricao) || bate(i.Codigo); })
    };
    var total = r.mp.length + r.p.length + r.sp.length + r.a.length + r.doc.length + r.gl.length + r.rp.length + r.t.length;
    function linha(href, cod, nome, extra) {
      return '<div class="doc-item"><i class="fas fa-arrow-right fa-stack-ico" aria-hidden="true"></i><div>' +
        '<div class="tit"><a href="' + href + '"><span class="cod">' + esc(cod) + '</span> ' + esc(nome) + '</a></div>' +
        (extra ? '<div class="meta">' + extra + '</div>' : '') + '</div></div>';
    }
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Busca' }]) +
      '<div class="pp-sec-h" style="margin-top:0"><h2>Resultados para “' + esc(q) + '”</h2><div class="linha" aria-hidden="true"></div></div>' +
      (total ? '' : '<p class="pp-vazio">Nada encontrado. Tente outro termo ou navegue pela lista de <a href="#/catalogo">processos</a>.</p>') +
      grupo('Macroprocessos', r.mp, function (m) { return linha('#/mp/' + encodeURIComponent(m.Codigo), m._cod || m.Codigo, m.Nome, esc(m.Categoria)); }) +
      grupo('Processos', r.p, function (p) { return linha('#/p/' + encodeURIComponent(p.Codigo), codDisp(p.Codigo), p.Nome, esc(p.Status_Mapeamento) + ' · ' + p.Percentual + '%'); }) +
      grupo('Subprocessos', r.sp, function (s) { return linha('#/sp/' + encodeURIComponent(s.Codigo), codDisp(s.Codigo), s.Nome, ''); }) +
      grupo('Atividades', r.a, function (a) { return linha('#/a/' + encodeURIComponent(a.Codigo), codDisp(a.Codigo), a.Nome, esc(a.Responsavel_Ator || '')); }) +
      grupo('Tarefas', r.t, function (t) { return linha('#/t/' + encodeURIComponent(t.Codigo), codDisp(t.Codigo), t.Nome, esc(t.Tipo_Tarefa || '')); }) +
      grupo('Documentos', r.doc, function (x) {
        return '<div class="doc-item"><i class="fas fa-file fa-stack-ico" aria-hidden="true"></i><div><div class="tit">' +
          (x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '</a>' : esc(x.Titulo)) +
          '</div><div class="meta">' + linkVinculos(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</div></div></div>';
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
    if ((c = $('#cntRepositorio'))) c.textContent = DADOS.repo.length;
    if ((c = $('#cntNugep'))) c.textContent = DADOS.nugep.length;
    if ((c = $('#cntGlossario'))) c.textContent = DADOS.glossario.length;
    if ((c = $('#cntFaq'))) c.textContent = DADOS.faq.length;
    ligarAcoesCabecalho();
    montarNotificacoes();
    if (window.PPUI) PPUI.setMenuSections([
      { rotulo: 'Início · Cadeia de Valor', href: '#/', icone: 'fa-house', meta: DADOS.macros.length + ' macro' },
      { rotulo: 'Catálogo de processos', href: '#/catalogo', icone: 'fa-layer-group', meta: DADOS.procs.length },
      { rotulo: 'Dashboard gerencial', href: '#/dashboard', icone: 'fa-chart-pie' },
      { rotulo: 'Repositório de materiais', href: '#/repositorio', icone: 'fa-toolbox' },
      { rotulo: 'Documentos', href: '#/documentos', icone: 'fa-folder-open', meta: DADOS.docs.length },
      { rotulo: 'Radar de riscos', href: '#/riscos', icone: 'fa-shield-halved', meta: DADOS.riscos.length },
      { rotulo: 'Indicadores', href: '#/indicadores', icone: 'fa-chart-line', meta: DADOS.inds.length },
      { rotulo: 'NUGEP', href: '#/nugep', icone: 'fa-people-group' },
      { rotulo: 'Glossário', href: '#/glossario', icone: 'fa-spell-check', meta: DADOS.glossario.length },
      { rotulo: 'Perguntas frequentes', href: '#/faq', icone: 'fa-circle-question', meta: DADOS.faq.length }
    ]);
    rota();
  }
  /* ── NOTIFICATION — central de alertas operacionais montada a partir
     dos dados reais: riscos críticos abertos e prazos vencidos ou
     próximos. Cada item traz tag de status, título, informação
     cronológica e conteúdo, separados por divider. ── */
  function montarNotificacoes() {
    var painelR = $('#notifPanelRiscos'), painelP = $('#notifPanelPrazos');
    if (!painelR || !painelP) return;
    var hoje = new Date().toISOString().slice(0, 10);
    var riscos = (DADOS.riscos || []).filter(function (r) {
      return (r._classe === 'Alto' || r._classe === 'Extremo') && !/encerrad/i.test(String(r.Status || ''));
    });
    var prazos = (DADOS.procs || []).filter(function (p) {
      return p.Prazo_Previsto && p.Prazo_Previsto < hoje && p._status !== 'concluido';
    });
    function item(tag, titulo, quando, texto, href) {
      return '<a class="br-item" href="' + href + '">' +
        '<span class="br-tag status small ' + tag + '" aria-hidden="true"></span>' +
        '<span class="text-bold">' + esc(titulo) + '</span>' +
        '<span class="text-medium">' + esc(quando) + '</span>' +
        '<span class="item-sub">' + esc(texto) + '</span></a>' +
        '<span class="br-divider" role="presentation"></span>';
    }
    painelR.innerHTML = riscos.length ? riscos.map(function (r) {
      return item('danger', r.Descricao_Risco || 'Risco crítico',
        'Nível ' + r._classe + ' · P×I ' + r._nivel,
        esc(r.ID) + ' · ' + (r.Status || 'Aberto'),
        '#/riscos');
    }).join('') : '<div class="empty-state"><i class="fas fa-circle-check" aria-hidden="true"></i>Nenhum risco crítico aberto.</div>';
    painelP.innerHTML = prazos.length ? prazos.map(function (p) {
      return item('warning', p.Nome, 'Prazo em ' + fmtData(p.Prazo_Previsto),
        codDisp(p.Codigo) + ' · ' + (p.Status_Mapeamento || '') + ' · ' + pctNorm(p.Percentual) + '%',
        '#/p/' + encodeURIComponent(p.Codigo));
    }).join('') : '<div class="empty-state"><i class="fas fa-circle-check" aria-hidden="true"></i>Nenhum mapeamento com prazo vencido.</div>';
    var total = riscos.length + prazos.length;
    var cR = $('#notifCountRiscos'), cP = $('#notifCountPrazos'), badge = $('#notifBadge');
    if (cR) cR.textContent = riscos.length;
    if (cP) cP.textContent = prazos.length;
    if (badge) { badge.textContent = total; badge.hidden = total === 0; }
  }

  function iniciar() {
    var v = $('#viewInicio');
    if (v) v.innerHTML = '<div class="pp-loading">' +
      '<div class="br-loading medium" role="progressbar" aria-label="Carregando dados do painel"></div>' +
      '<p class="loading-label" role="status" aria-live="polite">Carregando dados do painel…</p></div>';
    carregarDados().then(posCarga).catch(function (e) {
      console.error(e);
      if (v) v.innerHTML = '<div class="br-message warning" role="alert"><div class="icon"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i></div><div class="content"><span class="message-title">Não foi possível carregar os dados.</span> <span class="message-body">Verifique se data/painel-processos-dados.xlsx está publicado (ou gere js/dados.js com scripts/planilha_para_js.py). Detalhe: ' + esc(e.message) + '</span></div></div>';
    });
  }
  window.addEventListener('hashchange', rota);
  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
