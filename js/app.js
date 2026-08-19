/* ════════════════════════════════════════════════════════════════════
   REPOSITÓRIO DE PROCESSOS — aplicação (dados + rotas + telas).
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
    // Corrigido: faltavam 6 abas reais (Metricas/Medicoes substituíram
    // Indicadores; Papeis, Regras, Cultura_Processos e Iniciativas nunca
    // tinham entrado aqui) — carregarXlsx() só busca o que está nesta
    // lista, então elas SEMPRE voltavam vazias ao carregar da planilha real
    // ou do Google Sheets (só apareciam com dados porque js/dados.js
    // embutido cobria o buraco, escondendo o problema em qualquer preview).
    abas: ['Macroprocessos', 'Processos', 'Subprocessos', 'Atividades', 'Tarefas',
           'Documentos', 'Riscos', 'Metricas', 'Medicoes', 'Papeis', 'Regras',
           'Cultura_Processos', 'Iniciativas', 'Competencias',
           'Jornada', 'Repositorio', 'NUGEP', 'Equipe_Gerenciamento_Processos',
           'Glossario', 'FAQ', 'Siglas', 'Parametros']
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
  /* Data para exibição. A planilha guarda AAAA-MM-DD como texto, e é isso
     que o LEIA-ME manda preencher — mas o Excel converte o que parece data
     em número de série, e quem preenche à mão às vezes digita DD/MM/AAAA.
     As três formas são aceitas, para uma digitação fora do padrão não
     apagar a coluna inteira da tela. */
  function fmtData(v) {
    if (v === null || v === undefined || v === '') return '—';
    if (v instanceof Date) return dataBR(v.getDate(), v.getMonth() + 1, v.getFullYear());
    var t = String(v).trim();
    // Número de série do Excel (dias desde 30/12/1899).
    if (/^\d+(\.\d+)?$/.test(t) && +t > 20000 && +t < 80000) {
      var d = new Date(Date.UTC(1899, 11, 30) + Math.round(+t) * 86400000);
      return dataBR(d.getUTCDate(), d.getUTCMonth() + 1, d.getUTCFullYear());
    }
    var iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
    if (iso) return dataBR(+iso[3], +iso[2], +iso[1]);
    var br = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/.exec(t);
    if (br) return dataBR(+br[1], +br[2], +br[3]);
    return t;
  }
  function dataBR(d, m, a) {
    return (d < 10 ? '0' : '') + d + '/' + (m < 10 ? '0' : '') + m + '/' + a;
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
    return Math.round(n <= 1 ? n * 100 : n /* fração (0,45) ou percentual (45) — ver LEIA-ME */);
  }
  function simNao(v) {
    if (v === true) return true;
    return /^s/i.test(String(v == null ? '' : v).trim());
  }

  /* ── camada de dados ──────────────────────────────────────────────── */
  var DADOS = null;   // coleções normalizadas
  var IDX = null;     // índices por código / vínculo
  var FONTE = '';     // descrição da origem carregada
  // Métrica de uso do repositório (CBOK 4.2.2.3: "a métrica chave para um
  // bom repositório é a sua utilização") — site estático sem login nem
  // backend, então só é possível contar visitas NESTE navegador, não o
  // total de usuários da Companhia; o rótulo no dashboard é explícito
  // sobre essa limitação, para não sugerir uma métrica que não é.
  var VISITAS_NAVEGADOR = 0;
  function contarVisitaRepositorio() {
    try {
      var n = (parseInt(localStorage.getItem('pp_visitas_repositorio'), 10) || 0) + 1;
      localStorage.setItem('pp_visitas_repositorio', String(n));
      return n;
    } catch (e) { return 0; }
  }

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
    // Cache-bust: sem isso, o navegador pode reservir a planilha antiga do
    // cache HTTP mesmo depois do arquivo mudar no servidor.
    return fetch(CONFIG.arquivoXlsx + '?t=' + Date.now()).then(function (r) {
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
    /* Tira espaço nas pontas de todo texto vindo da planilha antes de qualquer
       outro processamento — um código ("MP-06 ") ou vínculo com espaço a mais
       nunca bate numa chave de índice, e o processo some da cadeia de valor
       sem erro nenhum no console. */
    function pega(aba) {
      return (bruto[aba] || []).map(function (l) {
        var o = {};
        Object.keys(l).forEach(function (k) { o[k] = typeof l[k] === 'string' ? l[k].trim() : l[k]; });
        return o;
      });
    }
    var dd = {
      macros: pega('Macroprocessos'),
      procs: pega('Processos'),
      subs: pega('Subprocessos'),
      ativs: pega('Atividades'),
      tarefas: pega('Tarefas'),
      docs: pega('Documentos'),
      riscos: pega('Riscos'),
      metricas: pega('Metricas'),
      medicoes: pega('Medicoes'),
      papeis: pega('Papeis'),
      regras: pega('Regras'),
      culturaProcessos: pega('Cultura_Processos'),
      iniciativas: pega('Iniciativas'),
      competencias: pega('Competencias'),
      jornada: pega('Jornada'),
      repo: pega('Repositorio'),
      nugep: pega('NUGEP'),
      equipeGestao: pega('Equipe_Gerenciamento_Processos'),
      glossario: pega('Glossario'),
      faq: pega('FAQ'),
      siglas: pega('Siglas'),
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
    // Medição (dado bruto), Métrica (valor calculado, com meta e polaridade)
    // e Indicador (a leitura simples dela, o selo "Meta atingida"/"Meta não
    // atingida") são três coisas diferentes no CBOK — cada Medição pertence
    // a uma Métrica (Metrica_ID); o "resultado atual" e a data de referência
    // da métrica vêm sempre da medição mais recente, nunca de uma célula
    // reescrita a cada apuração.
    dd.medicoes.forEach(function (x) { x.Data_Medicao = isoData(x.Data_Medicao); x.Valor = num(x.Valor); });
    var medicoesPorMetrica = {};
    dd.medicoes.forEach(function (x) { (medicoesPorMetrica[x.Metrica_ID] = medicoesPorMetrica[x.Metrica_ID] || []).push(x); });
    dd.metricas.forEach(function (x) {
      x.Meta = num(x.Meta);
      var hist = (medicoesPorMetrica[x.ID] || []).slice().sort(function (a, b) { return (b.Data_Medicao || '').localeCompare(a.Data_Medicao || ''); });
      x._medicoes = hist;
      x._ultima = hist[0] || null;
      x.Resultado_Atual = x._ultima ? x._ultima.Valor : null;
      x.Ultima_Medicao = x._ultima ? x._ultima.Data_Medicao : null;
      x._sit = situacaoInd(x);
      x._proxima = proximaMedicaoPrevista(x.Periodicidade, x.Ultima_Medicao);
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
    dd.equipeGestao.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    /* Foto e hierarquia dos integrantes: se a planilha em uso ainda não tem
       as colunas Foto e Hierarquia, aproveita o que está em js/dados.js,
       casando pelo e-mail — assim o avatar aparece com foto em qualquer
       fonte de dados, e a planilha passa a mandar assim que as colunas
       forem criadas (scripts/gerar_planilha.py já as gera). */
    (function () {
      var emb = (window.PAINEL_DADOS && window.PAINEL_DADOS.NUGEP) || [];
      if (!emb.length) return;
      var porEmail = {};
      emb.forEach(function (x) { if (x.Email) porEmail[String(x.Email).toLowerCase()] = x; });
      dd.nugep.forEach(function (m) {
        var ref = porEmail[String(m.Email || '').toLowerCase()];
        if (!ref) return;
        if (!m.Foto) m.Foto = ref.Foto || '';
        if (m.Hierarquia == null || m.Hierarquia === '') m.Hierarquia = ref.Hierarquia || 0;
      });
      // Chefias (níveis 1 e 2) ausentes da planilha entram pelo embutido —
      // elas existem no painel só para mostrar a hierarquia da unidade.
      emb.forEach(function (x) {
        var n = +(x.Hierarquia || 0);
        if (n !== 1 && n !== 2) return;
        var ja = dd.nugep.some(function (m) {
          return String(m.Email || '').toLowerCase() === String(x.Email || '').toLowerCase();
        });
        if (!ja) dd.nugep.push(Object.assign({}, x));
      });
      dd.nugep.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    })();
    dd.faq.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.glossario.sort(function (a, b) { return String(a.Termo || '').localeCompare(String(b.Termo || ''), 'pt-BR'); });
    dd.params = {};
    dd.parametros.forEach(function (x) { if (x.Chave) dd.params[x.Chave] = x.Valor || ''; });

    var idx = { mp: {}, p: {}, sp: {}, a: {}, t: {}, pByCod: {}, spByCod: {}, aByCod: {}, tByCod: {},
      procsPorMacro: {}, subsPorPai: {}, ativsPorPai: {}, tarefasPorAtiv: {},
      reusoPorAlvo: {}, vinc: { docs: {}, riscos: {}, metricas: {}, papeis: {}, regras: {} } };
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
    dd.procs.sort(function (a, b) { return String(a.Trilha).localeCompare(String(b.Trilha)); });
    dd.procs.forEach(function (p) {
      idx.p[p.Trilha] = p;
      if (!idx.pByCod[p.Codigo]) idx.pByCod[p.Codigo] = p;
      (idx.procsPorMacro[p.Macroprocesso] = idx.procsPorMacro[p.Macroprocesso] || []).push(p);
    });
    dd.subs.sort(function (a, b) { return (a.Ordem || 0) - (b.Ordem || 0); });
    dd.subs.forEach(function (s) {
      idx.sp[s.Trilha] = s;
      if (!idx.spByCod[s.Codigo]) idx.spByCod[s.Codigo] = s;
      // Vinculo_Pai aponta para um Processo (P-...) OU para outro Subprocesso (SP-...) —
      // o CBOK 4.0 não fixa a profundidade da decomposição ("Levels Vary in Number and
      // Name"): um subprocesso pode conter outro subprocesso, tantos níveis quanto o
      // processo exigir, até chegar à atividade. Agrupa pela TRILHA do pai — o código
      // bruto de Vinculo_Pai reinicia a cada vínculo e não identifica mais um único item.
      (idx.subsPorPai[trilhaPai(s.Trilha)] = idx.subsPorPai[trilhaPai(s.Trilha)] || []).push(s);
      // Subprocesso reutilizável (o "Call Activity" do BPMN 2.0 — no Bizagi,
      // que este painel já usa para os diagramas, o próprio elemento se
      // chama "Subprocesso Reutilizável"): mora nativamente sob Vinculo_Pai,
      // mas Reutilizado_Em lista o CAMINHO COMPLETO de outros processos/
      // subprocessos (ex. "MG-01/PP-03/SP-03") — de qualquer macroprocesso,
      // em qualquer nível de aninhamento — que também o chamam, sem duplicar
      // o mapeamento.
      if (s.Reutilizavel === 'Sim' && s.Reutilizado_Em) {
        listar(s.Reutilizado_Em).forEach(function (alvo) {
          var chave = normalizaCaminho(alvo);
          (idx.reusoPorAlvo[chave] = idx.reusoPorAlvo[chave] || []).push(s);
        });
      }
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
      idx.a[a.Trilha] = a;
      if (!idx.aByCod[a.Codigo]) idx.aByCod[a.Codigo] = a;
      (idx.ativsPorPai[trilhaPai(a.Trilha)] = idx.ativsPorPai[trilhaPai(a.Trilha)] || []).push(a);
    });
    dd.tarefas.sort(function (a, b) { return (num(a.Ordem) || 0) - (num(b.Ordem) || 0); });
    dd.tarefas.forEach(function (t) {
      idx.t[t.Trilha] = t;
      if (!idx.tByCod[t.Codigo]) idx.tByCod[t.Codigo] = t;
      (idx.tarefasPorAtiv[trilhaPai(t.Trilha)] = idx.tarefasPorAtiv[trilhaPai(t.Trilha)] || []).push(t);
    });
    function vincula(mapa, item) {
      listaVinculoPares(item.Vinculo_Nivel, item.Vinculo_Codigo).forEach(function (par) {
        var nivel = par[0], codigo = par[1], alvo = codigo;
        // Vinculo_Codigo agora pode vir como caminho completo ("MS-01/PP-01")
        // ou código isolado (formato antigo) — resolve para a Trilha real do
        // item, a mesma chave canônica que vinculados() usa na leitura.
        if (nivel !== 'Macroprocesso') { var r = resolveRef(nivel, codigo, idx); if (r) alvo = r.Trilha; }
        var ch = nivel + '|' + alvo;
        (mapa[ch] = mapa[ch] || []).push(item);
      });
    }
    dd.docs.forEach(function (x) { vincula(idx.vinc.docs, x); });
    dd.riscos.forEach(function (x) { vincula(idx.vinc.riscos, x); });
    dd.metricas.forEach(function (x) { vincula(idx.vinc.metricas, x); });
    dd.papeis.forEach(function (x) { vincula(idx.vinc.papeis, x); });
    dd.regras.forEach(function (x) { vincula(idx.vinc.regras, x); });
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
  // Próxima medição esperada = última medição + intervalo da periodicidade.
  function proximaMedicaoPrevista(periodicidade, dataUltima) {
    var meses = { 'Mensal': 1, 'Bimestral': 2, 'Trimestral': 3, 'Semestral': 6, 'Anual': 12 }[periodicidade];
    if (!meses || !dataUltima) return null;
    var d = new Date(dataUltima + 'T00:00:00');
    if (isNaN(d.getTime())) return null;
    d.setMonth(d.getMonth() + meses);
    return d.toISOString().slice(0, 10);
  }
  function situacaoInd(x) {
    if (x.Resultado_Atual == null) return 'Sem medição';
    if (x.Meta == null) return 'Sem meta';
    var maior = /^maior/i.test(String(x.Polaridade || ''));
    var ok = maior ? x.Resultado_Atual >= x.Meta : x.Resultado_Atual <= x.Meta;
    // Um único rótulo de "não atingida": "Acima/Abaixo da meta" lia como bom
    // fora do contexto de cor — um indicador "menor melhor" acima do alvo é
    // resultado ruim, mas o texto por si soava positivo. A direção (Meta vs
    // Resultado, já nas colunas ao lado, e Polaridade no detalhe) mostra o
    // "quanto" sem um rótulo ambíguo carregando o veredito.
    return ok ? 'Meta atingida' : 'Meta não atingida';
  }
  // N-000 (Norma de Gestão dos Instrumentos Normativos), item 4.2.3: todo
  // instrumento normativo deve ser revisado no máximo a cada 3 anos. Só se
  // aplica aos tipos que a Norma trata como instrumento normativo de fato
  // (Norma interna/Procedimento/Manual) — não a atas, diagramas, relatórios etc.
  var TIPOS_NORMATIVOS_N000 = { 'Norma interna': 1, 'Procedimento (PRO)': 1, 'Manual': 1 };
  function revisaoVencida(x) {
    if (!TIPOS_NORMATIVOS_N000[x.Tipo_Documento] || !x.Data || x.Situacao === 'Revogado') return false;
    var d = new Date(x.Data + 'T00:00:00');
    if (isNaN(d.getTime())) return false;
    d.setFullYear(d.getFullYear() + 3);
    return isoData(d) < hojeISO();
  }

  /* ── componentes reutilizáveis (HTML) ─────────────────────────────── */
  var ICONE_STATUS = {
    'nao-iniciado': 'fa-circle', 'em-andamento': 'fa-spinner',
    'concluido': 'fa-check-circle', 'suspenso': 'fa-pause-circle'
  };
  var _idTag = 0;
  function tagStatus(st) {
    var rot = st || 'Não iniciado', cls = slug(rot);
    var id = 'tg-st-' + (++_idTag);
    return '<span class="br-tag tag-status ' + cls + '" aria-describedby="' + id + '">' +
      '<i class="fas ' + (ICONE_STATUS[cls] || 'fa-circle') + '" aria-hidden="true"></i>' +
      '<span id="' + id + '">' + esc(rot) + '</span></span>';
  }
  function tagCat(cat) {
    return '<span class="tag-cat ' + slug(cat) + '">' + esc(cat || '') + '</span>';
  }
  // Ícones em escada: quanto mais grave, mais "cheia" a marca — o nível se
  // lê sem depender da cor (diretriz de acessibilidade da Tag).
  var ICONE_NIVEL = {
    'baixo': 'fa-circle-check', 'moderado': 'fa-circle-exclamation',
    'alto': 'fa-triangle-exclamation', 'extremo': 'fa-radiation'
  };
  function tagNivel(cl) {
    var cls = slug(cl), id = 'tg-nv-' + (++_idTag);
    return '<span class="br-tag nivel-tag nivel-' + cls + '" aria-describedby="' + id + '">' +
      '<i class="fas ' + (ICONE_NIVEL[cls] || 'fa-circle') + '" aria-hidden="true"></i>' +
      '<span id="' + id + '">' + esc(cl) + '</span></span>';
  }
  // Maturidade de processos (CBOK, escala tipo CMM) — preenchimento manual
  // na planilha; ícones em escada como nivel-tag, para não depender só de cor.
  var ICONE_MATURIDADE = {
    'inicial': 'fa-seedling', 'repetivel': 'fa-arrows-rotate', 'definido': 'fa-clipboard-check',
    'gerenciado': 'fa-gauge-high', 'otimizado': 'fa-star'
  };
  function tagMaturidade(nivel) {
    var cls = slug(nivel), id = 'tg-mt-' + (++_idTag);
    return '<span class="br-tag maturidade-tag maturidade-' + cls + '" aria-describedby="' + id + '">' +
      '<i class="fas ' + (ICONE_MATURIDADE[cls] || 'fa-circle') + '" aria-hidden="true"></i>' +
      '<span id="' + id + '">' + esc(nivel) + '</span></span>';
  }
  // Tipos de atividade (CBOK 3.1.7): agregação de valor, transferência
  // (handoff) e controle — definições completas no Glossário.
  var ICONE_TIPO_ATIV = { 'agregacao-de-valor': 'fa-gem', 'transferencia': 'fa-right-left', 'controle': 'fa-shield-halved' };
  function tagTipoAtividade(tipo) {
    var cls = slug(tipo), id = 'tg-ta-' + (++_idTag);
    return '<span class="br-tag tipo-ativ-tag tipo-ativ-' + cls + '" aria-describedby="' + id + '">' +
      '<i class="fas ' + (ICONE_TIPO_ATIV[cls] || 'fa-circle') + '" aria-hidden="true"></i>' +
      '<span id="' + id + '">' + termoLink(tipo) + '</span></span>';
  }
  function barraPct(p) {
    p = pctNorm(p);
    return '<div class="pct"><div class="trilho"><div class="barra" style="width:' + p +
      '%"></div></div><span class="valor">' + p + '%</span></div>';
  }
  // Sigla conhecida pela aba Siglas da planilha (Lista de Nomes e Siglas
  // oficial, Decisão da Presidência nº 601/2025) — carregada em DADOS.siglas,
  // não mais em código, para a UNP editar um nome sem tocar em nada aqui.
  var _siglasIdx = null;
  function siglasIdx() {
    if (!_siglasIdx) {
      _siglasIdx = {};
      (DADOS.siglas || []).forEach(function (u) { _siglasIdx[u.Sigla] = u.Nome; });
    }
    return _siglasIdx;
  }
  function siglaConhecida(codigo) {
    var idx = siglasIdx(), partes = String(codigo || '').trim().split('/'), acc = '';
    if (!partes[0]) return false;
    for (var i = 0; i < partes.length; i++) { acc = i === 0 ? partes[i] : acc + '/' + partes[i]; if (idx[acc]) return true; }
    return false;
  }
  // Sigla de unidade como link para a ficha dela no Glossário (Tooltip é
  // componente depreciado no DS — a referência abre a definição completa
  // em vez de um balão flutuante, sem the código não bater com nada conhecido.
  function siglaTag(codigo) {
    var cod = String(codigo == null ? '' : codigo).trim();
    if (!cod) return '';
    if (!siglaConhecida(cod)) return esc(cod);
    return '<a class="termo-link" href="#/glossario?aba=siglas&q=' + encodeURIComponent(cod) + '">' + esc(cod) + '</a>';
  }
  // Termo técnico (ex.: SIPOC) como link para a definição no Glossário.
  function termoLink(termo, textoVisivel) {
    return '<a class="termo-link" href="#/glossario?aba=termos&q=' + encodeURIComponent(termo) + '">' + esc(textoVisivel || termo) + '</a>';
  }
  function chips(str, icone, comoSigla) {
    var itens = listar(str);
    if (!itens.length) return '<span class="pp-vazio">Não informado</span>';
    return '<div class="chip-lista">' + itens.map(function (x) {
      return '<span class="chip">' + (icone ? '<i class="fas ' + icone + '" aria-hidden="true"></i> ' : '') + (comoSigla ? siglaTag(x) : esc(x)) + '</span>';
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
  // Lista simples (ul) ou ordenada (ol) a partir de "a; b; c" — para campos
  // com frases inteiras por item (passos, critérios), onde o chip-pill de
  // chips() ficaria estranho.
  function listaTexto(str, ordenada) {
    var itens = listar(str);
    if (!itens.length) return null;
    var tag = ordenada ? 'ol' : 'ul';
    return '<' + tag + '>' + itens.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</' + tag + '>';
  }
  function campo(rotulo, valorHtml, span2, categoria) {
    var cls = (span2 ? 'span2 ' : '') + (categoria ? 'campo-' + categoria : '');
    return '<div' + (cls.trim() ? ' class="' + cls.trim() + '"' : '') + '><dt>' + esc(rotulo) + '</dt><dd>' +
      (valorHtml || '<span class="pp-vazio">Não informado</span>') + '</dd></div>';
  }
  // Gestor(a) do processo: pessoa que responde pelo processo no dia a dia —
  // diferente do Ponto_Focal_Nugep (o contato durante o mapeamento/análise).
  // Ponto_Focal_Nugep guarda só o nome (ex. "Carlos Eduardo Lima (UNP)") — o
  // contato completo vem daqui, resgatado por nome na aba NUGEP.
  function pontoFocalNugepHtml(p) {
    var nome = String(p.Ponto_Focal_Nugep || '').trim();
    if (!nome) return '';
    var alvo = nome.toLowerCase();
    var m = (DADOS.nugep || []).filter(function (x) { return alvo.indexOf(String(x.Nome || '').toLowerCase()) === 0; })[0];
    if (!m) return esc(nome);
    return '<strong>' + esc(m.Nome) + '</strong>' + contatoNugep(m).replace('class="nugep-contato"', 'class="nugep-contato" style="border-bottom:0;padding-bottom:0"') +
      (m.Unidade_Sigla ? '<div style="margin-top:2px">' + siglaTag(m.Unidade_Sigla) + '</div>' : '');
  }
  function gestorProcessoHtml(p) {
    if (!p.Gestor_Nome) return '';
    var contatos = [];
    if (p.Gestor_Email) contatos.push('<a href="mailto:' + esc(p.Gestor_Email) + '"><i class="fas fa-envelope" aria-hidden="true"></i> ' + esc(p.Gestor_Email) + '</a>');
    if (p.Gestor_Telefone) contatos.push('<a href="tel:+55' + esc(String(p.Gestor_Telefone).replace(/\D/g, '')) + '"><i class="fas fa-phone" aria-hidden="true"></i> ' + esc(p.Gestor_Telefone) + '</a>');
    return '<strong>' + esc(p.Gestor_Nome) + '</strong>' +
      (contatos.length ? '<div class="nugep-contato" style="border-bottom:0;padding-bottom:0">' + contatos.join('') + '</div>' : '') +
      (p.Gestor_Unidade_Organica ? '<div style="margin-top:2px">' + siglaTag(p.Gestor_Unidade_Organica) + '</div>' : '');
  }
  // Sigla exibida de cada camada, sempre com DUAS letras, para que o nível
  // se leia no próprio código: MG/MF/MS (macroprocessos, por tipo), PP
  // (processo), SP (subprocesso), AT (atividade), TR (tarefa). O código da
  // planilha (MP-, P-, A-, T-) continua sendo a chave de vínculo — só a
  // exibição muda, então nenhuma planilha precisa ser regerada.
  // Trilha (ex.: "MS-01 › PP-01 › SP-03 › SP-04") é o identificador de
  // verdade de Processos/Subprocessos/Atividades/Tarefas: os códigos
  // (PP-/SP-/AT-/TR-) reiniciam em 01 a cada novo vínculo com o pai e não
  // são mais únicos sozinhos. trilhaPai tira o último segmento — a trilha
  // do pai direto, seja ele um Processo ou outro Subprocesso.
  function trilhaPai(trilha) {
    var partes = String(trilha || '').split(' › ');
    partes.pop();
    return partes.join(' › ');
  }
  function ultimoSegmento(trilha) {
    var partes = String(trilha || '').split(' › ');
    return partes[partes.length - 1] || '';
  }
  // Reutilizado_Em passou a guardar o CAMINHO COMPLETO (aceita "/" ou "›"
  // como separador) — normaliza para o mesmo formato de Trilha, para
  // comparar/indexar os dois com a mesma chave.
  function normalizaCaminho(s) {
    return String(s || '').split(/[\/›]/).map(function (x) { return x.trim(); }).filter(Boolean).join(' › ');
  }
  function codDisp(c) {
    var s = String(c == null ? '' : c);
    if (IDX && IDX.mp[s]) return IDX.mp[s]._cod || s;
    if (s.indexOf('P-') === 0) return 'PP-' + s.slice(2);
    if (s.indexOf('A-') === 0) return 'AT-' + s.slice(2);
    if (s.indexOf('T-') === 0) return 'TR-' + s.slice(2);
    return s;                                // SP- e MG/MF/MS já têm 2 letras
  }
  // Trilha completa para exibição: mesma normalização de separador que
  // normalizaCaminho (aceita "/" ou "›" na origem), com cada segmento
  // passando por codDisp — cobre também um vínculo que guarde só o código
  // isolado em vez do caminho completo.
  function trilhaDisp(trilha) {
    return String(trilha || '').split(/\s*[\/›]\s*/).filter(Boolean).map(codDisp).join(' › ');
  }
  // Nível de um código/trilha, lido do PRÓPRIO texto (último segmento) — mais
  // confiável que o Vinculo_Nivel gravado quando a célula lista códigos de
  // níveis diferentes (ex. "MS-01 › PP-01; MS-01 › PP-01 › SP-03"), caso em
  // que um só Vinculo_Nivel não consegue valer para os dois ao mesmo tempo.
  function nivelDoCodigo(codigo) {
    var s = String(codigo || '');
    if (/TR-/.test(s)) return 'Tarefa';
    if (/AT-/.test(s)) return 'Atividade';
    if (/SP-/.test(s)) return 'Subprocesso';
    if (/PP-/.test(s)) return 'Processo';
    if (/^\s*(MG|MF|MS)-/.test(s)) return 'Macroprocesso';
    return '';
  }
  var NIVEL_PREFIXO = { 'Macroprocesso': 'mp', 'Processo': 'p', 'Subprocesso': 'sp',
    'Atividade': 'a', 'Tarefa': 't' };
  var NIVEL_ROTULO = { 'Macroprocesso': 'Macroprocesso', 'Processo': 'Processo',
    'Subprocesso': 'Subprocesso', 'Atividade': 'Atividade', 'Tarefa': 'Tarefa' };
  function nivelRotulo(n) { return NIVEL_ROTULO[n] || n; }
  // Resolve um código BRUTO de vínculo (Vinculo_Codigo, Processos_Relacionados)
  // para o item real. Aceita os dois formatos: caminho completo ("MS-01/PP-01"
  // ou "MS-01 › PP-01", já sem ambiguidade — normaliza e busca por Trilha) OU
  // código isolado (formato antigo, ambíguo — cai no primeiro item daquele
  // código, via ByCod, para planilhas ainda não migradas não pararem de achar
  // nada).
  function resolveRef(nivel, codigo, idxAtual) {
    var base = idxAtual || IDX;
    var porCod = nivel === 'Processo' ? base.pByCod : nivel === 'Subprocesso' ? base.spByCod
      : nivel === 'Atividade' ? base.aByCod : nivel === 'Tarefa' ? base.tByCod : null;
    if (!porCod) return null;
    var s = String(codigo || '');
    if (s.indexOf('/') >= 0 || s.indexOf('›') >= 0) {
      var trilha = normalizaCaminho(s);
      var porTrilha = nivel === 'Processo' ? base.p[trilha] : nivel === 'Subprocesso' ? base.sp[trilha]
        : nivel === 'Atividade' ? base.a[trilha] : base.t[trilha];
      if (porTrilha) return porTrilha;
    }
    return porCod[ultimoSegmento(s)] || porCod[s] || null;
  }
  function rotaDe(nivel, codigo) {
    var pre = NIVEL_PREFIXO[nivel];
    if (!pre) return '#/';
    var item = resolveRef(nivel, codigo);
    return '#/' + pre + '/' + encodeURIComponent(item ? item.Trilha : codigo);
  }
  /* Categoria (Gerencial/Finalístico/Suporte) do macroprocesso ao qual o
     item pertence, subindo a hierarquia a partir de qualquer nível. */
  function categoriaDe(nivel, codigo) {
    var mp = null;
    if (nivel === 'Macroprocesso') mp = IDX.mp[codigo];
    else if (nivel === 'Processo') { var p = resolveRef('Processo', codigo); mp = p && IDX.mp[p.Macroprocesso]; }
    else if (nivel === 'Subprocesso') { var s0 = resolveRef('Subprocesso', codigo); var pp = s0 && processoDoSubprocesso(s0.Trilha); mp = pp && IDX.mp[pp.Macroprocesso]; }
    else if (nivel === 'Atividade') {
      var a = resolveRef('Atividade', codigo), an = a && ancestraisDaAtividade(a); mp = an && an.mp;
    } else if (nivel === 'Tarefa') {
      var t = resolveRef('Tarefa', codigo), a2 = t && resolveRef('Atividade', t.Atividade);
      var an2 = a2 && ancestraisDaAtividade(a2); mp = an2 && an2.mp;
    }
    return mp && mp.Categoria ? mp.Categoria : '';
  }
  // Macroprocesso/Subprocesso/Atividade/Tarefa têm definição própria no
  // Glossário ("Processo" não — só aparece dentro de outras definições,
  // sem entrada exata, então fica sem link para não abrir uma busca vaga).
  var NIVEIS_NO_GLOSSARIO = { 'Macroprocesso': 1, 'Subprocesso': 1, 'Atividade': 1, 'Tarefa': 1 };
  function eyebrowFicha(nivel, codigo) {
    var cat = categoriaDe(nivel, codigo);
    var rotulo = NIVEIS_NO_GLOSSARIO[nivel] ? termoLink(nivel) : nivel;
    return '<span class="eyebrow">' + rotulo + (cat ? ' · ' + esc(cat) : '') + '</span>';
  }
  function nomeDe(nivel, codigo) {
    var it = nivel === 'Macroprocesso' ? IDX.mp[codigo] : resolveRef(nivel, codigo);
    return it ? it.Nome : codigo;
  }
  function linkVinculo(nivel, codigo) {
    return '<a href="' + rotaDe(nivel, codigo) + '"><span class="cod">' + esc(trilhaDisp(codigo)) +
      // A sigla de duas letras no início do link já diz o nível (MG/MF/MS,
      // PP, SP, AT, TR) — o rótulo por extesso ao lado era redundante e
      // alargava a coluna sem informar nada de novo.
      '</span> ' + esc(nomeDe(nivel, codigo)) + '</a>';
  }
  // Um mesmo item (documento, risco ou indicador) pode estar vinculado a dois ou mais
  // processos, dois ou mais subprocessos, ou a um processo e um subprocesso ao mesmo
  // tempo: Vinculo_Nivel e Vinculo_Codigo aceitam listas paralelas separadas por ';'.
  function listaVinculoPares(nivelStr, codigoStr) {
    var niveis = listar(nivelStr), codigos = listar(codigoStr), pares = [];
    // Um só Vinculo_Nivel com vários códigos em Vinculo_Codigo é o caso comum
    // (mesmo nível repetido) — repete o nível para não descartar os códigos
    // além do primeiro.
    if (niveis.length === 1 && codigos.length > 1) niveis = codigos.map(function () { return niveis[0]; });
    for (var i = 0; i < Math.max(niveis.length, codigos.length); i++) {
      // O nível de CADA código vem do próprio código, não da célula — cobre o
      // caso raro de vincular a itens de níveis diferentes na mesma linha.
      var niv = (codigos[i] && nivelDoCodigo(codigos[i])) || niveis[i];
      if (niv && codigos[i]) pares.push([niv, codigos[i]]);
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
      '" aria-label="Página Inicial" title="Página Inicial"><span class="sr-only">Página Inicial</span>' +
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
     Cada marco lê exatamente uma coluna real da aba Processos, na mesma
     ordem e com os mesmos dez nomes da Metodologia (RES 031/2025): Conhecer
     o processo, Processo modelado, Subprocessos modelados, AS-IS modelado,
     AS-IS validado, Procedimento aprovado, Processo publicado, TO-BE
     elaborado, TO-BE aprovado, Processo transformado. */
  var MARCOS = [
    { rot: 'Conhecer o processo', campos: ['M1_Conhecer_Processo'],
      desc: 'Primeira reunião com a área para conhecer o processo: apresentar a metodologia, entender o contexto e coletar o formulário de levantamento.' },
    { rot: 'Processo modelado', campos: ['M2_Processo_Modelado'],
      desc: 'Oficinas de modelagem do macroprocesso e do processo em BPMN, com a equipe de mapeamento e a área.' },
    { rot: 'Subprocessos modelados', campos: ['M3_Subprocessos_Modelados'],
      desc: 'Oficinas de modelagem dos subprocessos — podendo aprofundar vários níveis (subprocesso dentro de subprocesso) — inclusive identificando subprocessos ainda não mapeados.' },
    { rot: 'AS-IS modelado', campos: ['M4_ASIS_Modelado'],
      desc: 'O conjunto de diagramas AS-IS (macroprocesso, processo, subprocessos e atividades) está consolidado.' },
    { rot: 'AS-IS validado', campos: ['M5_ASIS_Validado'],
      desc: 'O dono do processo validou formalmente o AS-IS como retrato da realidade atual.' },
    { rot: 'Procedimento aprovado', campos: ['M6_Procedimento_Aprovado'],
      desc: 'O procedimento (PRO) foi aprovado pela Diretoria Executiva (DEX) — pronto para orientar a execução do processo.' },
    { rot: 'Processo publicado', campos: ['M7_Processo_Publicado'],
      desc: 'O processo, seus diagramas e o PRO foram publicados no repositório institucional, disponíveis para consulta de toda a Empresa.' },
    { rot: 'TO-BE elaborado', campos: ['M8_TOBE_Elaborado'],
      desc: 'O redesenho (TO-BE) do processo foi elaborado, com melhorias, riscos residuais e indicadores propostos.' },
    { rot: 'TO-BE aprovado', campos: ['M9_TOBE_Aprovado'],
      desc: 'O dono do processo validou o TO-BE, que foi aprovado pela Diretoria Executiva (DEX).' },
    { rot: 'Processo transformado', campos: ['M10_Processo_Transformado'],
      desc: 'O redesenho aprovado foi implantado: o processo passou a ser executado na forma TO-BE, com os ganhos acompanhados pelos indicadores.' }
  ];
  var MARCOS_ROTULOS = MARCOS.map(function (m) { return m.rot; });
  var MARCOS_CAMPOS = MARCOS.map(function (m) { return m.campos[0]; });
  var MARCOS_DESCRICOES = MARCOS.map(function (m) { return m.desc; });
  // Valor do marco i no processo p — primeira coluna preenchida da lista.
  /* Tooltip padrão gov.br (Componente Tooltip) no lugar do title nativo do
     navegador: um ícone discreto de informação ("i") abre o balão no hover
     E no foco, o que torna o texto alcançável por teclado e leitor de tela
     — o title nativo não é. Uso: rótulo + dica('texto explicativo'). */
  /* Estado vazio no padrão Dados Ausentes (gov.br DS > Padrões).
     Anatomia: Apoio Visual (ilustração, opcional) + Título (condicional,
     obrigatório quando o vazio ocupa a tela ou um bloco grande) +
     Mensagem (obrigatória) + Suporte para Ações (opcional, ao final).
     O padrão pede tom neutro — nunca sugerir que o usuário errou — e
     mensagem específica, dizendo o critério usado e o que fazer a seguir.
     Ilustrações do pacote oficial; a de "empty-space" é a narrativa que
     corresponde à ausência de conteúdo. */
  function vazio(titulo, mensagem, opts) {
    opts = opts || {};
    var acoes = opts.acoes || [];
    return '<div class="pp-vazio pp-vazio-ilus" role="status">' +
      '<img src="img/ilustracoes/' + (opts.img || 'empty-space/empty-space-07.png') + '" alt="" aria-hidden="true">' +
      '<div class="pv-txt">' +
      (titulo ? '<p class="pv-titulo">' + titulo + '</p>' : '') +
      '<p class="pv-msg">' + mensagem + '</p></div>' +
      (acoes.length ? '<div class="pv-acoes">' + acoes.map(function (a, i) {
        // Hierarquia do componente Button: ênfase secundária na ação
        // principal, terciária nas demais (regra do padrão).
        var cls = 'br-button small' + (i === 0 ? ' secondary' : '');
        var rot = (a.icone ? '<i class="fas ' + a.icone + '" aria-hidden="true"></i> ' : '') + esc(a.rotulo);
        return a.href ? '<a class="' + cls + '" href="' + a.href + '">' + rot + '</a>'
          : '<button class="' + cls + '" type="button" data-vazio-acao="' + a.acao + '">' + rot + '</button>';
      }).join('') + '</div>' : '') +
      '</div>';
  }
  function dica(texto, place) {
    if (!texto) return '';
    return '<span class="tooltip-wrap dica"><button type="button" class="dica-btn" data-tooltip-trigger' +
      ' aria-label="Mais informações"><i class="fas fa-info-circle" aria-hidden="true"></i></button>' +
      '<span class="br-tooltip small" role="tooltip" info place="' + (place || 'top') + '">' +
      '<span class="subtext">' + esc(texto) + '</span></span></span>';
  }
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
    if (/^em\s*andamento$/i.test(s)) return 'andamento';
    return '';
  }
  /* Marcos do mapeamento pelo Componente Step (tipo complexo, indicador
     numérico, orientação horizontal): a jornada tem ordem lógica linear,
     que é exatamente o caso de uso do componente. Os dez marcos ficam
     todos à vista — sem data-scroll, portanto sem arrastar barra na
     horizontal; abaixo de 992px a trilha vira vertical (variante
     documentada do componente) em vez de espremer dez colunas.
     Os indicadores não são interativos (o painel só informa o progresso),
     por isso vêm com disabled; a etapa atual é o primeiro marco ainda
     pendente e "não se aplica" recebe alerta em cor de atenção. A
     descrição de cada marco sai do title nativo e vira Tooltip gov.br,
     acionado pelo "i" ao lado do rótulo (o gatilho é o container, já que
     o botão desabilitado não recebe eventos de mouse). */
  function marcosHtml(p) {
    var est = MARCOS.map(function (mk, i) { return marcoEstado(valMarco(p, i)); });
    var atual = -1;
    for (var k = 0; k < est.length; k++) { if (est[k] === '') { atual = k; break; } }
    var total = MARCOS.length;
    return '<nav class="br-step marcos-step" data-label="bottom" role="none">' +
      '<div class="step-progress" role="listbox" aria-orientation="horizontal" aria-label="Marcos do mapeamento">' +
      MARCOS.map(function (mk, i) {
        var e = est[i], feito = e === 'feito', na = e === 'na', andamento = e === 'andamento', ativo = i === atual;
        var rotulo = 'M' + (i + 1) + ' — ' + MARCOS_ROTULOS[i] +
          (feito ? ' (concluído)' : na ? ' (não se aplica)' : andamento ? ' (em andamento — meia nota)' : ativo ? ' (em andamento)' : ' (pendente)');
        var texto = MARCOS_DESCRICOES[i] + (na ? ' — Não se aplica a este processo.' : '');
        return '<span class="tooltip-wrap step-wrap" data-tooltip-trigger tabindex="0">' +
          '<button class="step-progress-btn' + (feito ? ' is-done' : andamento ? ' is-andamento' : '') + '" type="button" role="option"' +
          ' step-num="' + (i + 1) + '" aria-posinset="' + (i + 1) + '" aria-setsize="' + total + '"' +
          ' aria-selected="' + (ativo ? 'true' : 'false') + '"' + (ativo ? ' active' : '') + ' disabled' +
          (na ? ' data-alert="warning"' : '') +
          ' aria-label="' + esc(rotulo) + '">' +
          '<span class="step-info">' + esc(MARCOS_ROTULOS[i]) +
          ' <i class="fas fa-info-circle dica-i" aria-hidden="true"></i></span>' +
          (na ? '<span class="step-alert"></span>' : '') +
          '</button>' +
          '<span class="br-tooltip small" role="tooltip" info place="bottom">' +
          '<span class="text">' + esc(rotulo) + '</span>' +
          '<span class="subtext">' + esc(texto) + '</span></span></span>';
      }).join('') +
      '</div></nav>';
  }
  // Sobe a cadeia de um subprocesso até achar seu Processo — como o CBOK permite
  // subprocesso dentro de subprocesso (profundidade variável), o "pai" de um
  // subprocesso pode ser outro subprocesso (código "SP-...") em vez de um
  // processo (código "P-...") direto. Retorna [maisFundo, ..., maisRaso],
  // sempre terminando no subprocesso que é filho direto do Processo.
  function cadeiaSubprocessos(spTrilha) {
    var cadeia = [], atual = IDX.sp[spTrilha], guarda = {};
    while (atual && !guarda[atual.Trilha]) {
      cadeia.push(atual); guarda[atual.Trilha] = true;
      if (String(atual.Vinculo_Pai || '').indexOf('SP-') === 0) atual = IDX.sp[trilhaPai(atual.Trilha)];
      else break;
    }
    return cadeia;
  }
  function processoDoSubprocesso(spTrilha) {
    var cadeia = cadeiaSubprocessos(spTrilha);
    var raiz = cadeia[cadeia.length - 1];
    return raiz ? IDX.p[trilhaPai(raiz.Trilha)] : null;
  }
  function ehCodigoSub(c) { return String(c || '').indexOf('SP-') === 0; }
  function atividadesDe(trilha) { return IDX.ativsPorPai[trilha] || []; }
  // Pai direto de uma atividade: um Subprocesso (SP-...) ou o próprio Processo
  // (P-...), quando o processo não tem subprocessos.
  function paiDaAtividade(a) {
    if (!a) return null;
    var c = a._pai || a.Vinculo_Pai || a.Subprocesso || '';
    var pt = trilhaPai(a.Trilha);
    if (ehCodigoSub(c)) return IDX.sp[pt] ? { tipo: 'sp', item: IDX.sp[pt] } : null;
    return IDX.p[pt] ? { tipo: 'p', item: IDX.p[pt] } : null;
  }
  // Ancestrais de uma atividade, prontos para o breadcrumb e para o card
  // "Navegar para": macroprocesso, processo e a cadeia de subprocessos — que
  // vem vazia quando a atividade pende direto do processo.
  function ancestraisDaAtividade(a) {
    var pai = paiDaAtividade(a), sp = null, p = null;
    if (pai && pai.tipo === 'sp') { sp = pai.item; p = processoDoSubprocesso(sp.Trilha); }
    else if (pai) { p = pai.item; }
    return { pai: pai, sp: sp, p: p, mp: p ? IDX.mp[p.Macroprocesso] : null,
      cadeiaSp: sp ? cadeiaSubprocessos(sp.Trilha).slice().reverse() : [] };
  }
  // Filhos de codigoPai para fins de agregação: os subprocessos nativos MAIS
  // os subprocessos reutilizáveis chamados aqui (definidos em outro lugar,
  // mas que executam de verdade quando este processo roda). O parâmetro
  // "vistos" evita loop infinito se um cadastro criar um ciclo de reúso.
  function subsReutilizadosEm(codigo) { return IDX.reusoPorAlvo[normalizaCaminho(codigo)] || []; }
  function filhosSubDe(codigoPai) { return (IDX.subsPorPai[codigoPai] || []).concat(subsReutilizadosEm(codigoPai)); }
  function contarAtividadesRecursivo(codigoPai, vistos) {
    vistos = vistos || {};
    if (vistos[codigoPai]) return 0;
    vistos[codigoPai] = true;
    // atividades ligadas DIRETO a codigoPai (processo sem subprocesso, ou
    // subprocesso) + as de todos os subprocessos descendentes, em qualquer
    // profundidade (subprocesso dentro de subprocesso) e incluindo
    // subprocessos reutilizáveis chamados a partir daqui.
    var total = atividadesDe(codigoPai).length;
    filhosSubDe(codigoPai).forEach(function (s) {
      total += contarAtividadesRecursivo(s.Trilha, vistos);
    });
    return total;
  }
  function contarSubprocessosRecursivo(codigoPai, vistos) {
    vistos = vistos || {};
    if (vistos[codigoPai]) return 0;
    vistos[codigoPai] = true;
    var subs = filhosSubDe(codigoPai), total = subs.length;
    subs.forEach(function (s) { total += contarSubprocessosRecursivo(s.Trilha, vistos); });
    return total;
  }
  function contarTarefasRecursivo(codigoPai, vistos) {
    vistos = vistos || {};
    if (vistos[codigoPai]) return 0;
    vistos[codigoPai] = true;
    var total = 0;
    atividadesDe(codigoPai).forEach(function (a) { total += (IDX.tarefasPorAtiv[a.Trilha] || []).length; });
    filhosSubDe(codigoPai).forEach(function (s) { total += contarTarefasRecursivo(s.Trilha, vistos); });
    return total;
  }
  // Duração: hora útil é a unidade de medida da coluna Duracao_Estimada da
  // aba Tarefas. A duração de uma atividade é o somatório das suas tarefas;
  // a de um processo (ou subprocesso) é o somatório recursivo — o caminho
  // crítico do processo, não o caminho feliz (o cenário mais otimista).
  function horasTarefa(t) { var h = Number(t.Duracao_Estimada); return isFinite(h) && h > 0 ? h : 0; }
  function duracaoAtividadeHoras(a) {
    return (IDX.tarefasPorAtiv[a.Trilha] || []).reduce(function (soma, t) { return soma + horasTarefa(t); }, 0);
  }
  function duracaoRecursivaHoras(codigoPai, vistos) {
    vistos = vistos || {};
    if (vistos[codigoPai]) return 0;
    vistos[codigoPai] = true;
    var soma = atividadesDe(codigoPai).reduce(function (s, a) { return s + duracaoAtividadeHoras(a); }, 0);
    filhosSubDe(codigoPai).forEach(function (s) { soma += duracaoRecursivaHoras(s.Trilha, vistos); });
    return soma;
  }
  function formatarHorasUteis(horas) {
    var h = Math.round(horas * 10) / 10;
    return plural(h, 'hora', 'horas');
  }
  // Card de um subprocesso na lista "Subprocessos vinculados": o mesmo
  // cartão, com uma marcação extra quando o item é reutilizado aqui (mora
  // nativamente em outro processo/subprocesso, de outro macroprocesso ou não).
  function subCardHtml(s, marcarReuso) {
    var home = String(s.Vinculo_Pai || '');
    var homeTrilha = trilhaPai(s.Trilha);
    var origem = ehCodigoSub(home) ? IDX.sp[homeTrilha] : IDX.p[homeTrilha];
    var origemTxt = origem ? origem.Trilha + ' — ' + origem.Nome : home;
    return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="#/sp/' + encodeURIComponent(s.Trilha) + '"><div class="topo"><div><span class="cod">' + esc(trilhaDisp(s.Trilha)) + '</span>' +
      '<div class="nome" style="font-size:var(--fs-sm)">' + esc(s.Nome) + '</div>' +
      (marcarReuso ? '<div class="reuso-def"><span class="br-tag info small"><i class="fas fa-link" aria-hidden="true"></i> Subprocesso reutilizável</span>' +
        '<span class="reuso-origem">Definido em ' + esc(origemTxt) + '</span></div>' : '') +
      '</div></div></a>';
  }
  // Mesma marcação de reúso, numa linha de tabelaGov (usada em "Subprocessos
  // deste subprocesso", que já é tabela, não cartão).
  function linhaSubprocesso(sf, marcarReuso) {
    return '<tr data-link><td class="cod">' + esc(codDisp(sf.Codigo)) + '</td><td><a href="#/sp/' + encodeURIComponent(sf.Trilha) + '"><strong>' + esc(sf.Nome) + '</strong></a>' +
      (marcarReuso ? '<div style="margin-top:4px"><span class="br-tag info small"><i class="fas fa-link" aria-hidden="true"></i> Subprocesso reutilizável</span></div>' : '') +
      (sf.Descricao ? '<div class="pp-muted" style="font-size:var(--fs-sm)">' + esc(sf.Descricao) + '</div>' : '') + '</td>' +
      '<td><a class="br-button secondary small" href="#/sp/' + encodeURIComponent(sf.Trilha) + '">Abrir ficha</a></td></tr>' +
      detalheLinha([
        ['Entradas', listar(sf.Entradas).map(esc).join('; ')],
        ['Saídas', listar(sf.Saidas).map(esc).join('; ')],
        ['Sistemas', listar(sf.Sistemas).map(esc).join('; ')],
        ['Unidade responsável', esc(sf.Unidade_Organica_Responsavel || '')]
      ]);
  }
  // Todo lugar que chama um subprocesso reutilizável: o pai nativo
  // (Vinculo_Pai) mais cada código em Reutilizado_Em, sem repetir.
  function usosDoReutilizavel(s) {
    var alvos = [trilhaPai(s.Trilha)].concat(listar(s.Reutilizado_Em).map(normalizaCaminho));
    var vistos = {};
    return alvos.filter(function (tr) { if (!tr || vistos[tr]) return false; vistos[tr] = true; return true; }).map(function (tr) {
      var ultimo = ultimoSegmento(tr), ehSub = ehCodigoSub(ultimo), item = ehSub ? IDX.sp[tr] : IDX.p[tr];
      return { cod: codDisp(item ? item.Codigo : ultimo), nome: item ? item.Nome : tr, trilha: item ? item.Trilha : tr, href: ehSub ? '#/sp/' + encodeURIComponent(tr) : '#/p/' + encodeURIComponent(tr) };
    });
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
          // Padrão Content Overflow > Truncamento: todo texto truncado
          // precisa dar acesso ao conteúdo completo. Aqui o title cumpre o
          // papel (o link já é alcançável por teclado, e o leitor de tela
          // lê o texto inteiro, que não é cortado no DOM).
          '<span class="hier-nome" title="' + esc(codDisp(it.codigo) + ' — ' + it.nome) + '">' +
          esc(codDisp(it.codigo)) + ' — ' + esc(it.nome) + '</span></span>' +
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
      '<p class="pp-muted" style="font-size:var(--fs-sm);margin-top:6px"><i class="fas fa-info-circle" aria-hidden="true"></i> Diagrama interativo (Bizagi Web Publish). Se o quadro acima aparecer em branco, o servidor pode estar bloqueando a incorporação ou exigir rede interna — use o link abaixo.</p>' +
      '<div class="diagrama-acoes"><a class="br-button secondary small" href="' + href +
      '" target="_blank" rel="noopener"><i class="fas fa-external-link-alt" aria-hidden="true"></i>&nbsp;Abrir em Nova Aba<span class="sr-only"> (abre em nova aba)</span></a></div>';
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
        '<div class="tit">' + tit + (revisaoVencida(x) ? ' <span class="br-tag revisao-vencida"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> Revisão vencida</span>' : '') + '</div><div class="meta">' + esc(x.Tipo_Documento || 'Documento') +
        (x.Versao ? ' · v' + esc(x.Versao) : '') + (x.Data ? ' · ' + fmtData(x.Data) : '') +
        (x.Situacao ? ' · ' + esc(x.Situacao) : '') + (x.Ato_Aprovacao ? ' · ' + esc(x.Ato_Aprovacao) : '') + '</div></div></div>';
    }).join('');
  }
  /* ── TABLE (br-table) ── Envelope padrão gov.br para qualquer conteúdo
     tabular do painel. Monta a anatomia completa em volta das colunas
     recebidas: Barra de Título com título e ações utilitárias (menu de
     densidade + gatilho de busca), Barra de Busca que cobre o título
     enquanto ativa, Área de Dados com <caption>, cabeçalhos ordenáveis
     (aria-sort + ícones sort/sort-up/sort-down) e wrapper responsivo.
     O comportamento (ordenar, buscar, trocar densidade) é genérico e
     opera sobre o DOM — ver BRTableInit em govbr-ui.js. ── */
  var tblSeq = 0;
  function tabelaGov(cfg) {
    var id = cfg.id || ('tbl' + (++tblSeq));
    var dens = cfg.densidade || 'small';
    var cols = cfg.colunas;
    return '<div class="br-table ' + dens + '" id="' + id + '" data-search="data-search"' +
      ' data-selection="data-selection" data-generic="data-generic">' +
      '<div class="table-header"><div class="top-bar">' +
      '<div class="table-title">' + esc(cfg.titulo) + '</div>' +
      '<div class="actions-trigger text-nowrap">' +
      '<span class="tooltip-wrap">' +
      '<button class="br-button circle small" type="button" id="' + id + '-dens" data-tooltip-trigger data-toggle="dropdown" data-target="' + id + '-densMenu" aria-label="Definir densidade da tabela" aria-haspopup="true" aria-expanded="false"><i class="fas fa-ellipsis-v" aria-hidden="true"></i></button>' +
      '<span class="br-tooltip small" role="tooltip" info place="bottom">' +
      '<span class="text">Densidade da tabela</span>' +
      '<span class="subtext">Compacta, regular ou espaçada</span></span>' +
      '</span>' +
      '<div class="br-list dd-target" id="' + id + '-densMenu" role="menu" aria-labelledby="' + id + '-dens" hidden>' +
      [['small', 'Compacta'], ['medium', 'Regular'], ['large', 'Espaçada']].map(function (o, i) {
        return (i ? '<span class="br-divider" role="presentation"></span>' : '') +
          '<button class="br-item" type="button" role="menuitem" data-density="' + o[0] + '"' +
          (dens === o[0] ? ' aria-current="true"' : '') + '>' + o[1] + '</button>';
      }).join('') + '</div></div>' +
      '<div class="search-trigger"><button class="br-button circle small" type="button" data-toggle="search" aria-label="Abrir busca" aria-controls="' + id + '-busca" aria-expanded="false"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '</div>' +
      '<div class="search-bar"><div class="br-input">' +
      '<label for="' + id + '-busca">Pesquisar na tabela</label>' +
      '<input id="' + id + '-busca" type="search" placeholder="Pesquisar na tabela">' +
      '<button class="br-button circle small" type="button" aria-label="Pesquisar"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '<button class="br-button circle small" type="button" data-dismiss="search" aria-label="Fechar pesquisa"><i class="fas fa-times" aria-hidden="true"></i></button>' +
      '</div></div>' +
      '<div class="responsive"><table><caption>' + esc(cfg.titulo) + '</caption><colgroup>' +
      cols.map(function (c) {
        // min: largura mínima em px. O .col-curta encolhe a coluna ao
        // conteúdo (width:1%), o que espremia códigos e rótulos curtos
        // (R-001, "Extremo", "Em tratamento") em duas ou três linhas.
        return '<col' + (c.principal ? ' class="col-principal"' : c.curta ? ' class="col-curta"' : '') +
          (c.min ? ' style="min-width:' + c.min + 'px"' : '') + '>';
      }).join('') + '</colgroup><thead><tr>' +
      cols.map(function (c) {
        if (c.fixa) return '<th scope="col">' + esc(c.r) + '</th>';
        return '<th scope="col"><button type="button" class="sort-btn" data-sort aria-label="Ordenar por ' + esc(c.r) + '">' +
          esc(c.r) + '<i class="fas fa-sort" aria-hidden="true"></i></button></th>';
      }).join('') + '</tr></thead><tbody>' + cfg.linhas + '</tbody></table></div>' +
      '<div class="table-footer"><span class="tbl-total"><strong>' + cfg.total + '</strong> ' + esc(cfg.rotuloTotal || 'registros') + '</span></div>' +
      '</div>';
  }

  /* Linha de detalhe (comportamento 3 do Componente Table). A informação
     secundária sai das colunas e vem para cá: a tabela fica com o que o
     usuário precisa ver de relance, e o resto continua a um clique. Uma
     informação aparece OU na coluna OU no detalhe, nunca nos dois. */
  function detalheLinha(pares) {
    var itens = pares.filter(function (p) { return p[1] && String(p[1]).trim() && p[1] !== '—'; });
    if (!itens.length) return '';
    return '<tr class="collapse"><td><dl class="tbl-detalhe">' +
      itens.map(function (p) {
        return '<div><dt>' + esc(p[0]) + '</dt><dd>' + p[1] + '</dd></div>';
      }).join('') + '</dl></td></tr>';
  }

  // Tabela de atividades — usada na ficha do Subprocesso e na ficha do
  // Processo (quando as atividades penduram direto no processo).
  // unidadeSigla vem do PAI (subprocesso ou processo) sendo exibido, não de
  // um campo próprio da atividade — todas as atividades de uma mesma
  // tabela pertencem ao mesmo pai, logo à mesma unidade responsável.
  function tabelaAtividadesHtml(ativs, vazio, unidadeSigla) {
    if (!ativs.length) return '<p class="pp-vazio">' + (vazio || 'Nenhuma atividade cadastrada.') + '</p>';
    return tabelaGov({
      titulo: 'Atividades', rotuloTotal: ativs.length === 1 ? 'atividade' : 'atividades', total: ativs.length,
      colunas: [{ r: 'Código', curta: true, min: 132 }, { r: 'Atividade', principal: true },
        { r: 'Unidade Orgânica Responsável' }, { r: 'Duração estimada', curta: true }],
      linhas: ativs.map(function (a) {
        var nt = (IDX.tarefasPorAtiv[a.Trilha] || []).length;
        return '<tr data-link><td class="cod">' + esc(codDisp(a.Codigo)) + '</td>' +
          '<td><a href="#/a/' + encodeURIComponent(a.Trilha) + '"><strong>' + esc(a.Nome) + '</strong></a></td>' +
          '<td>' + (unidadeSigla ? siglaTag(unidadeSigla) : '—') + '</td>' +
          '<td class="text-nowrap">' + (nt ? formatarHorasUteis(duracaoAtividadeHoras(a)) : '—') + '</td></tr>' +
          detalheLinha([
            ['Entradas', listar(a.Entradas).map(esc).join('; ')],
            ['Saídas', listar(a.Saidas).map(esc).join('; ')],
            ['Sistemas', listar(a.Sistemas).map(esc).join('; ')],
            ['Tarefas', nt ? plural(nt, 'tarefa', 'tarefas') : '']
          ]);
      }).join('')
    });
  }
  function ligarLinhasTabela() {
    $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        // A linha é atalho para a ficha, menos onde há controle próprio:
        // link, caixa de seleção, botão de expandir.
        if (ev.target.closest('a, input, label, button, .column-checkbox')) return;
        var lk = tr.querySelector('a'); if (lk) location.hash = lk.getAttribute('href');
      });
    });
  }
  function tabelaRiscosHtml(riscos, comVinculo) {
    if (!riscos.length) return '<p class="pp-vazio">Nenhum risco registrado.</p>';
    var cols = [{ r: 'ID', curta: true, min: 92 }, { r: 'Risco', principal: true }];
    if (comVinculo) cols.push({ r: 'Vinculado a', min: 200 });
    cols = cols.concat([{ r: 'P', curta: true }, { r: 'I', curta: true },
      { r: 'P×I', curta: true }, { r: 'Nível', curta: true, min: 128 },
      { r: 'Status', curta: true, min: 124 }]);
    return tabelaGov({
      titulo: 'Riscos', rotuloTotal: riscos.length === 1 ? 'risco' : 'riscos', total: riscos.length,
      colunas: cols,
      linhas: riscos.map(function (r) {
        return '<tr id="risco-' + esc(r.ID) + '"><td class="cod">' + esc(r.ID) + '</td>' +
          '<td data-th="Risco"><strong>' + esc(r.Descricao_Risco) + '</strong></td>' +
          (comVinculo ? '<td data-th="Vinculado a">' + linkVinculos(r.Vinculo_Nivel, r.Vinculo_Codigo) + '</td>' : '') +
          '<td class="num">' + r.Probabilidade_1a5 + '</td><td class="num">' + r.Impacto_1a5 +
          '</td><td class="num"><strong>' + r._nivel + '</strong></td><td>' + tagNivel(r._classe) +
          '</td><td>' + esc(r.Status || '—') + '</td></tr>' +
          detalheLinha([
            ['Tratamento e controles', esc(r.Controles_Tratamento || '')],
            ['Resposta ao risco', esc(r.Resposta || '')],
            ['Categoria', esc(r.Categoria || '')],
            ['Responsável', esc(r.Responsavel || '')],
            ['Cronograma do risco', esc(r.Cronograma_Risco || '')],
            ['Fatores (o que pode acionar o risco)', esc(r.Fatores || '')],
            ['Identificado em', fmtData(r.Data_Identificacao)],
            ['Prazo do tratamento', r.Prazo_Tratamento ? fmtData(r.Prazo_Tratamento) : '']
          ]);
      }).join('')
    });
  }
  function tabelaIndsHtml(inds, comVinculo) {
    if (!inds.length) return '<p class="pp-vazio">Nenhum indicador vinculado.</p>';
    var cols = [{ r: 'ID', curta: true, min: 100 }, { r: 'Indicador', principal: true }];
    if (comVinculo) cols.push({ r: 'Vinculado a', min: 200 });
    cols = cols.concat([{ r: 'Meta', curta: true }, { r: 'Resultado', curta: true }, { r: 'Situação', curta: true },
      { r: 'Última medição', curta: true }]);
    return tabelaGov({
      titulo: 'Indicadores', rotuloTotal: inds.length === 1 ? 'indicador' : 'indicadores', total: inds.length,
      colunas: cols,
      linhas: inds.map(function (x) {
        var cls = x._sit === 'Meta atingida' ? 'sit-ok' : (x._sit === 'Sem medição' || x._sit === 'Sem meta') ? 'sit-neutra' : 'sit-ruim';
        var un = x.Unidade ? ' ' + esc(x.Unidade) : '';
        return '<tr><td class="cod">' + esc(x.ID) + '</td>' +
          '<td><strong>' + esc(x.Nome) + '</strong></td>' + (comVinculo ? '<td>' + linkVinculos(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' : '') +
          '<td>' + (x.Meta != null ? x.Meta + un : '—') + '</td><td>' + (x.Resultado_Atual != null ? x.Resultado_Atual + un : '—') +
          '</td><td class="' + cls + '">' + esc(x._sit) + '</td>' +
          '<td>' + fmtData(x.Ultima_Medicao) + '</td></tr>' +
          detalheLinha([
            ['Categoria da métrica', esc(x.Categoria || '')],
            ['Fórmula de cálculo', esc(x.Descricao_Formula || '')],
            ['Critérios de desempenho', esc(x.Criterios_Desempenho || '')],
            ['Periodicidade', esc(x.Periodicidade || '')],
            ['Polaridade', esc(x.Polaridade || '')],
            ['Fonte', esc(x.Fonte || '')],
            ['Observação', esc(x.Observacoes || '')],
            ['Próxima medição prevista', x._proxima ? fmtData(x._proxima) + (x._proxima < hojeISO() ? ' <span class="sit-ruim">(atrasada)</span>' : '') : ''],
            // Medição é o dado bruto (CBOK) — o histórico completo mora na
            // aba Medicoes; aqui mostramos só as últimas, para contexto.
            ['Últimas medições', x._medicoes && x._medicoes.length
              ? x._medicoes.slice(0, 6).map(function (me) { return fmtData(me.Data_Medicao) + ': ' + esc(me.Valor) + (x.Unidade ? ' ' + esc(x.Unidade) : ''); }).join(' &nbsp;·&nbsp; ')
              : '']
          ]);
      }).join('')
    });
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
      '<i class="fas fa-caret-down" aria-hidden="true"></i></button></div>' +
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
  /* ── Barra de seções responsiva ─────────────────────────────────────
     As ações de página (alertas, compartilhar, imprimir) são âncoras fixas
     do cabeçalho e precisam estar sempre visíveis. Quando a linha aperta,
     as abas que não cabem migram para o menu "Mais" — em vez de rolarem
     para fora da tela, onde o próprio botão "Mais" desaparecia. ── */
  function ajustarAbas() {
    var nav = $('#navigation .tab-nav');
    var lista = nav && nav.querySelector('ul');
    if (!lista) return;
    var maisLi = $('#tab-mais') && $('#tab-mais').closest('.tab-item');
    var menu = $('#maisMenu');
    if (!maisLi || !menu) return;
    var itens = $all('li.tab-item', lista).filter(function (li) { return li !== maisLi; });
    // Reinicia: todas as abas visíveis e o menu sem itens migrados.
    itens.forEach(function (li) { li.hidden = false; });
    $all('[data-migrado]', menu).forEach(function (n) { n.remove(); });
    maisLi.hidden = false;

    var disponivel = nav.clientWidth;
    var larguraMais = maisLi.offsetWidth;
    var larguras = itens.map(function (li) { return li.offsetWidth; });
    var total = larguras.reduce(function (a, b) { return a + b; }, 0);
    var excedentes = [];
    if (total > disponivel) {
      // A seção em que o usuário está permanece sempre visível na faixa —
      // migrar a aba ativa esconderia o contexto atual.
      // A folga de 8px garante que o botão "Mais" caiba inteiro na faixa:
      // sem ela, o arredondamento das larguras deixava a seta do botão
      // encostada no limite do contêiner, que tem overflow escondido.
      var limite = disponivel - larguraMais - 8;
      for (var i = itens.length - 1; i >= 0 && total > limite; i--) {
        if (itens[i].classList.contains('active')) continue;
        total -= larguras[i];
        itens[i].hidden = true;
        excedentes.unshift(itens[i]);
      }
    }
    if (!excedentes.length) return;
    // Cada aba migrada vira um item da List do menu, preservando ícone,
    // rótulo e contador — sem duplicar ids.
    var frag = d.createDocumentFragment();
    excedentes.forEach(function (li) {
      var b = li.querySelector('[data-rota]');
      if (!b) return;
      var ic = b.querySelector('.name > i');
      var rot = b.querySelector('.tab-label-full');
      var cnt = b.querySelector('.tab-count');
      var item = d.createElement('button');
      item.type = 'button';
      item.className = 'br-item';
      item.setAttribute('data-migrado', '');
      item.setAttribute('data-rota', b.getAttribute('data-rota'));
      item.setAttribute('data-painel', b.getAttribute('data-painel'));
      item.innerHTML = '<i class="' + (ic ? ic.className : 'fas fa-circle') + '" aria-hidden="true"></i>' +
        '<span>' + esc(rot ? rot.textContent : b.textContent.trim()) + '</span>' +
        (cnt ? '<span class="tab-count">' + esc(cnt.textContent) + '</span>' : '');
      frag.appendChild(item);
    });
    var sep = d.createElement('span');
    sep.className = 'br-divider';
    sep.setAttribute('data-migrado', '');
    sep.setAttribute('aria-hidden', 'true');
    frag.appendChild(sep);
    var cabecalho = menu.querySelector('.header');
    menu.insertBefore(frag, cabecalho ? cabecalho.nextSibling : menu.firstChild);
    ajustando = true;
    mostrarPainel(painelAtual);
    ajustando = false;
  }
  var ajustando = false;
  var painelAtual = 'inicio';
  function mostrarPainel(id) {
    painelAtual = id;
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
    // Se a seção recém-aberta estava migrada para o "Mais", refaz o recorte
    // para trazê-la de volta à faixa.
    if (!ajustando) {
      var ativoLi = $('#navigation .tab-nav li.tab-item.active');
      if (ativoLi && ativoLi.hidden) ajustarAbas();
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
    else if ((m = h.match(/^#\/glossario(?:\?(.*))?$/))) {
      if (m[1]) {
        var qp = {}; m[1].split('&').forEach(function (par) { var kv = par.split('='); qp[kv[0]] = decodeURIComponent(kv[1] || ''); });
        if (qp.aba) filtroGloss.aba = qp.aba;
        if (qp.q != null) { if (filtroGloss.aba === 'siglas') filtroSiglas.q = qp.q; else filtroGloss.q = qp.q; }
      }
      renderGlossario(); mostrarPainel('glossario');
    }
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
    // Os tooltips do conteúdo recém-renderizado precisam ser ligados a cada
    // troca de tela (a inicialização é idempotente: ignora os já prontos).
    if (window.BRTooltipInit) window.BRTooltipInit();
    if (window.BRMessageInit) window.BRMessageInit();
    if (window.BRTableInit) window.BRTableInit();
    if (window.BRLoadingInit) window.BRLoadingInit();
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
      var cab = ['Codigo', 'Macroprocesso', 'Nome', 'Status_Mapeamento', 'Percentual',
        'Unidade_Organica_Responsavel', 'Ponto_Focal_Nugep', 'Prazo_Previsto'];
      // O CSV leva a sigla de exibição (PP-/AT-/TR-), a mesma que aparece na
      // tela — quem exporta compara com o painel, não com a planilha.
      var codigos = { Codigo: 1, Macroprocesso: 1, Vinculo_Pai: 1 };
      var linhas = [cab.join(';')].concat(DADOS.procs.map(function (p) {
        return cab.map(function (k) {
          var v = codigos[k] ? codDisp(p[k]) : p[k];
          return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
        }).join(';');
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
  var MP_ICONES = { 'MG-01': 'fa-bullseye', 'MG-02': 'fa-shield-halved', 'MF-01': 'fa-seedling',
    'MF-02': 'fa-droplet', 'MF-03': 'fa-water', 'MS-01': 'fa-file-contract',
    'MS-02': 'fa-users', 'MS-03': 'fa-laptop-code' };
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
  /* Texto vindo da aba Parâmetros, com o texto atual como reserva. */
  function par(chave, padrao) {
    var v = (DADOS.params || {})[chave];
    return (v == null || String(v).trim() === '') ? padrao : String(v);
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
      '<h1>' + esc(par('Titulo_Inicio', 'Mapeamento de processos da Codevasf')) + '</h1>' +
      (par('Subtitulo_Inicio', '') ? '<p class="pp-hero-sub">' + esc(par('Subtitulo_Inicio', '')) + '</p>' : '') +
      '</section>' +
      '<div class="kpi-grid">' +
      '<div class="kpi"><span class="num">' + DADOS.macros.length + '</span><span class="lbl">Macroprocessos' + dica('Total de macroprocessos na cadeia de valor institucional (gerenciais, finalísticos e de suporte).') + '</span><span class="sub">' + DADOS.subs.length + ' subprocessos · ' + DADOS.ativs.length + ' atividades · ' + DADOS.tarefas.length + ' tarefas</span></div>' +
      '<div class="kpi"><span class="num">' + procs.length + '</span><span class="lbl">Processos identificados' + dica('Total de processos identificados no portfólio de processos da Companhia.') + '</span><span class="sub">no portfólio atual</span></div>' +
      '<div class="kpi ok"><span class="num">' + concl + '</span><span class="lbl">Mapeamentos concluídos' + dica('Processos com todos os marcos do mapeamento (M1–M10) concluídos.') + '</span><span class="sub">' + andamento + ' em andamento</span></div>' +
      '<div class="kpi"><span class="num">' + media + '%</span><span class="lbl">Avanço médio' + dica('Percentual médio de execução do mapeamento entre todos os processos do portfólio.') + '</span><span class="sub">do mapeamento do portfólio</span></div>' +
      '<div class="kpi ' + (criticos ? 'erro' : 'ok') + '"><span class="num">' + criticos + '</span><span class="lbl">Riscos críticos abertos' + dica('Riscos classificados como Alto ou Extremo, ainda não encerrados.') + '</span><span class="sub">nível Alto ou Extremo</span></div>' +
      '</div>' +
      '<section class="pp-sec" id="sec-cadeia"><div class="pp-sec-h"><h2>Cadeia de Valor Integrada</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="cadeia">' +
      '<aside class="cv-aside cv-missao"><h3><i class="fas fa-flag" aria-hidden="true"></i> Missão</h3><p>' + esc(INSTITUCIONAL.missao) + '</p><h3><i class="fas fa-eye" aria-hidden="true"></i> Visão</h3><p>' + esc(INSTITUCIONAL.visao) + '</p></aside>' +
      '<div class="cv-centro">' + blocoCadeia('Macroprocessos Gerenciais', 'cat-gerencial', 'fa-compass', ger) +
      blocoCadeia('Macroprocessos Finalísticos — entrega de valor à sociedade', 'cat-finalistico', 'fa-hand-holding-heart', fin) + '</div>' +
      '<aside class="cv-aside cv-proposito"><h3><i class="fas fa-hand-holding-heart" aria-hidden="true"></i> Propósito</h3><p>' + esc(INSTITUCIONAL.proposito) + '</p></aside>' +
      '<div class="cv-suporte">' + blocoCadeia('Macroprocessos de Suporte', 'cat-suporte', 'fa-gears', sup) + '</div>' +
      '<div class="cv-valores"><strong><i class="fas fa-gem" aria-hidden="true"></i> Valores</strong>' + INSTITUCIONAL.valores.map(function (v) { return '<span class="cv-chip">' + esc(v) + '</span>'; }).join('') + '</div>' +
      '</div></section>' +
      cardCulturaProcessos();
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
      ' aria-controls="' + id + '-list" aria-expanded="false"><i class="fas fa-caret-down" aria-hidden="true"></i></button></div>' +
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
      '<ul class="checkbox-list horizontal" role="group" aria-labelledby="fStatusLabel">' +
      '<li><div class="br-checkbox"><input id="fStatus-todos" name="fStatus-todos" type="checkbox" data-parent="status"' +
      (todos || marcados ? ' checked' : '') + (marcados && !todos ? ' indeterminate' : '') + '>' +
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
    // checked + indeterminate juntos: é o que o utilitário Checkgroup
    // define para o estado intermediário.
    pai.checked = marcados > 0;
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
    return '<a class="proc-card" href="#/p/' + encodeURIComponent(p.Trilha) + '">' +
      '<div class="topo"><div><span class="cod" style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(trilhaDisp(p.Trilha)) + '</span>' +
      '<div class="nome">' + esc(p.Nome) + '</div></div>' + tagStatus(p.Status_Mapeamento) + '</div>' +
      '<div class="pp-muted" style="font-size:var(--fs-sm);margin-top:4px">' +
      [p.Unidade_Organica_Responsavel || '', marcoRotulo(marcoAtual(p))].filter(Boolean).map(esc).join(' · ') + '</div>' +
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
      (lista.length ? '<div class="proc-grid">' + pagFatia('catalogo', ordenarCat(lista), 10).map(cardProcesso).join('') + '</div>' +
        paginacaoHtml('catalogo', lista.length, 'processos', [10, 20, 40, 80])
        : vazio('Nenhum processo com esses filtros',
          'Ajuste os critérios ou limpe os filtros para ver todo o portfólio.',
          { acoes: [{ rotulo: 'Limpar Filtros', icone: 'fa-rotate-left', acao: 'cat-limpar' }] }));
  }
  function renderResultados() {
    var box = $('#catResultados');
    if (!box) return renderCatalogo();
    box.innerHTML = catResultadosHtml(catFiltrada());
    ligarPaginacao(box, renderResultados);
    // O rodapé do painel vive fora de #catResultados; a visibilidade do
    // botão "Limpar filtros" acompanha cada mudança de filtro.
    var rodape = $('#fRodapeFiltros');
    if (rodape) {
      rodape.hidden = !catTemFiltro();
      var ativos = rodape.querySelector('.filtros-ativos');
      var html = filtrosAtivosHtml();
      if (ativos) ativos.outerHTML = html || '<div class="filtros-ativos" hidden></div>';
      else rodape.insertAdjacentHTML('afterbegin', html);
    }
  }
  /* Filtros ativos como Tags de interação dispensáveis (Componente Tag,
     tipo interação): cada tag mostra um recorte em vigor e o botão fechar
     remove só aquele recorte. O rótulo tem id próprio, referenciado pelo
     aria-describedby do botão, como pede o exemplo oficial. */
  function filtrosAtivosHtml() {
    var tags = [];
    function tag(tipo, valor, icone, rotulo) {
      var id = 'ftag-' + tipo + '-' + String(valor).replace(/[^a-z0-9]+/gi, '-');
      tags.push('<span class="br-tag interaction small">' +
        '<i class="fas ' + icone + '" aria-hidden="true"></i>' +
        '<span id="' + id + '">' + esc(rotulo) + '</span>' +
        '<button class="br-button" type="button" aria-label="Remover filtro" aria-describedby="' + id + '"' +
        ' data-filtro-tipo="' + tipo + '" data-filtro-valor="' + esc(String(valor)) + '">' +
        '<i class="fas fa-times" aria-hidden="true"></i></button></span>');
    }
    filtroCat.macro.forEach(function (c) {
      var m = DADOS.macros.filter(function (x) { return x.Codigo === c; })[0];
      tag('macro', c, 'fa-diagram-project', m ? (m._cod || m.Codigo) : c);
    });
    filtroCat.status.forEach(function (v) {
      var r = STATUS_MAPEAMENTO.filter(function (s) { return slug(s) === v; })[0] || v;
      tag('status', v, 'fa-check-circle', r);
    });
    if (filtroCat.marco) tag('marco', filtroCat.marco, 'fa-flag', 'M' + filtroCat.marco);
    if (filtroCat.q) tag('q', 'q', 'fa-search', '“' + filtroCat.q + '”');
    if (filtroCat.de || filtroCat.ate) tag('prazo', 'prazo', 'fa-calendar-alt', periodoTexto());
    return tags.length ? '<div class="filtros-ativos" aria-label="Filtros ativos">' + tags.join('') + '</div>' : '';
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
      '<div class="pp-sec-h" style="margin-top:0"><h1>Portfólio de processos</h1><div class="linha" aria-hidden="true"></div></div>' +
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
      '<label for="fBusca">Pesquisar no portfólio</label>' +
      '<div class="input-group"><div class="input-icon"><i class="fas fa-search" aria-hidden="true"></i></div>' +
      '<input type="search" id="fBusca" placeholder="Código ou nome do processo" value="' + esc(filtroCat.q) + '">' +
      '</div></div>' +
      '</div>' +
      '<span class="br-divider" role="presentation"></span>' +
      '<div class="filtros-opcoes">' + filtroStatusHtml(DADOS.procs) + filtroOrdemHtml() + '</div>' +
      '<div class="filtros-rodape" id="fRodapeFiltros"' +
      (catTemFiltro() ? '' : ' hidden') + '>' + filtrosAtivosHtml() +
      '<button class="br-button secondary small" type="button" id="fLimparTudo"><i class="fas fa-rotate-left" aria-hidden="true"></i> Limpar Filtros</button></div>' +
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
    var rodape = $('#fRodapeFiltros');
    if (rodape) rodape.onclick = function (ev) {
      var b = ev.target.closest('[data-filtro-tipo]');
      if (!b) return;
      var t = b.getAttribute('data-filtro-tipo'), v = b.getAttribute('data-filtro-valor');
      if (t === 'macro') filtroCat.macro = filtroCat.macro.filter(function (x) { return x !== v; });
      else if (t === 'status') filtroCat.status = filtroCat.status.filter(function (x) { return x !== v; });
      else if (t === 'marco') filtroCat.marco = '';
      else if (t === 'q') filtroCat.q = '';
      else if (t === 'prazo') { filtroCat.de = ''; filtroCat.ate = ''; }
      PAG.catalogo.pag = 1;
      renderCatalogo();
    };
    function limparFiltrosCat() {
      filtroCat.macro = []; filtroCat.status = []; filtroCat.marco = '';
      filtroCat.q = ''; filtroCat.de = ''; filtroCat.ate = '';
      PAG.catalogo.pag = 1;
      renderCatalogo();
    }
    var limpar = $('#fLimparTudo');
    if (limpar) limpar.onclick = limparFiltrosCat;
    /* O botão "Limpar filtros" do estado vazio é redesenhado a cada
       filtragem, então a escuta fica no container (delegação). O container
       sobrevive aos redesenhos, daí a marca para não empilhar escutas. */
    if (!el.dataset.vazioLigado) {
      el.dataset.vazioLigado = '1';
      el.addEventListener('click', function (ev) {
        if (ev.target.closest('[data-vazio-acao="cat-limpar"]')) limparFiltrosCat();
      });
    }
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
    var inds = semIndicadores ? [] : vinculados('metricas', nivel, codigo);
    var papeis = vinculados('papeis', nivel, codigo);
    var regras = vinculados('regras', nivel, codigo);
    return (semIndicadores ? '' : '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> Indicadores de desempenho</h3>' + tabelaIndsHtml(inds, false) + '</div>') +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Riscos</h3>' + tabelaRiscosHtml(riscos, false) + '</div>' +
      (papeis.length ? '<div class="pp-card"><h3><i class="fas fa-users" aria-hidden="true"></i> Papéis e envolvidos</h3>' + listaPapeisHtml(papeis) + '</div>' : '') +
      (regras.length ? '<div class="pp-card"><h3><i class="fas fa-gavel" aria-hidden="true"></i> Regras de negócio</h3>' + listaRegrasHtml(regras) + '</div>' : '') +
      '<div class="pp-card"><h3><i class="fas fa-folder-open" aria-hidden="true"></i> Normativos e documentos vinculados</h3>' + listaDocsHtml(docs) + '</div>';
  }
  // RACI traduzido: Executa (R), Aprova (A), Consultado (C), Informado (I).
  function listaPapeisHtml(papeis) {
    if (!papeis.length) return '<p class="pp-vazio">Nenhum papel cadastrado.</p>';
    return '<div class="br-list raci-lista" role="list">' + papeis.map(function (r) {
      return '<div class="br-item raci-item" role="listitem"><span class="br-tag small raci-' + slug(r.Envolvimento || '') + '">' + esc(r.Envolvimento || '—') + '</span>' +
        '<div class="raci-tx"><strong>' + esc(r.Papel) + '</strong>' + (r.Unidade_Pessoa ? '<span class="pp-muted"> · ' + esc(r.Unidade_Pessoa) + '</span>' : '') + '</div></div>';
    }).join('') + '</div>';
  }
  function listaRegrasHtml(regras) {
    if (!regras.length) return '<p class="pp-vazio">Nenhuma regra cadastrada.</p>';
    return '<div class="br-list regras-lista" role="list">' + regras.map(function (r) {
      return '<div class="br-item regra-item" role="listitem"><div class="regra-topo"><span class="br-tag info small">' + esc(r.Tipo_Regra || '—') + '</span>' +
        (r.Fonte_Normativa ? '<span class="pp-muted">' + esc(r.Fonte_Normativa) + '</span>' : '') + '</div>' +
        '<div class="regra-nome">' + esc(r.Nome) + '</div>' +
        (r.Descricao ? '<p class="regra-desc">' + esc(r.Descricao) + '</p>' : '') + '</div>';
    }).join('') + '</div>';
  }
  // Cultura de Processos (CBOK 9.5.6) é autoavaliação ORGANIZACIONAL, não por
  // processo — fica no dashboard gerencial. Vem "Não avaliado" por padrão:
  // só a equipe pode dizer se cada característica já é realidade.
  function cardCulturaProcessos() {
    var itens = DADOS.culturaProcessos || [];
    if (!itens.length) return '';
    var sim = itens.filter(function (c) { return c.Situacao === 'Sim'; }).length;
    return '<div class="pp-card" style="margin-top:var(--sp5)"><h3><i class="fas fa-people-group" aria-hidden="true"></i> Cultura de processos</h3>' +
      '<div class="cultura-lista">' + itens.map(function (c) {
        return '<div class="cultura-item"><span class="br-tag small cultura-' + slug(c.Situacao || 'nao-avaliado') + '">' + esc(c.Situacao || 'Não avaliado') + '</span><span>' + esc(c.Caracteristica) + '</span></div>';
      }).join('') + '</div></div>';
  }
  // Iniciativas viabilizadas pelo repósitorio: uma das 3 métricas de um bom
  // repositório no CBOK (4.2.11) — fica vazio de propósito até a equipe
  // registrar um uso real; não é dado para eu inventar.
  function cardIniciativas() {
    var itens = DADOS.iniciativas || [];
    return '<div class="pp-card"><h3><i class="fas fa-rocket" aria-hidden="true"></i> Iniciativas viabilizadas pelo repositório</h3>' +
      (itens.length ? '<div class="br-list" role="list">' + itens.map(function (i) {
        return '<div class="br-item" role="listitem"><div><strong>' + esc(i.Titulo) + '</strong><span class="pp-muted"> · ' + fmtData(i.Data) + '</span>' +
          (i.Tipo ? ' <span class="br-tag info small">' + esc(i.Tipo) + '</span>' : '') + '</div>' +
          (i.Descricao ? '<p style="margin:2px 0 0">' + esc(i.Descricao) + '</p>' : '') + '</div>';
      }).join('') + '</div>' :
      '<p class="pp-vazio">Nenhuma iniciativa registrada ainda. Uma das três métricas de um bom repositório (CBOK 4.2.11) é quantas iniciativas ele viabiliza — registre na aba Iniciativas da planilha quando o repositório embasar uma decisão, automação ou projeto de melhoria.</p>') + '</div>';
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
        eyebrowFicha('Macroprocesso', m.Codigo) + '<h1>' + esc(m._cod || m.Codigo) + ' — ' + esc(m.Nome) + '</h1>' +
        '<div class="meta"><span>' + filhos.length + ' processos vinculados</span><span>· gerenciamento atual em ' + media + '%</span></div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do macroprocesso</h3><dl class="ficha-dl">' +
        campo('Objetivo', m.Objetivo && esc(m.Objetivo), false, 'desc') + campo('Descrição', m.Descricao && esc(m.Descricao), false, 'desc') +
        campo('Unidade Orgânica responsável', m.Unidade_Organica_Responsavel && siglaTag(m.Unidade_Organica_Responsavel), false, 'quem') +
        campo('Unidades Orgânicas Corresponsáveis', chips(m.Unidades_Organicas_Corresponsaveis, 'fa-people-group', true), false, 'quem') +
        campo('Entregas (produtos/serviços)', chips(m.Entregas), true, 'valor') +
        campo('Beneficiários', chips(m.Beneficiarios), false, 'valor') +
        campo('Partes interessadas', chips(m.Partes_Interessadas), false, 'valor') +
        campo('Sistemas utilizados', chips(m.Sistemas, 'fa-desktop'), false, 'tecnico') +
        (m.Observacoes ? campo('Observações', esc(m.Observacoes), true) : '') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama</h3>' + diagramaHtml(m.Imagem_Bizagi, m.Nome) + '</div>' +
        secVinculos('Macroprocesso', cod) +
        '</div><aside>' +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Processos vinculados</h3>' +
        (filhos.length ? filhos.map(function (p) {
          return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="#/p/' + encodeURIComponent(p.Trilha) + '">' +
            '<div class="topo"><div><span style="font-family:var(--noto-mono,monospace);font-size:var(--fs-sm);color:var(--gray-60)">' + esc(trilhaDisp(p.Trilha)) + '</span>' +
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
      var reusados = subsReutilizadosEm(cod);
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
        eyebrowFicha('Processo', p.Codigo) +
        '<h1>' + esc(codDisp(p.Codigo)) + ' — ' + esc(p.Nome) + '</h1>' +
        '<div class="meta">' + tagStatus(p.Status_Mapeamento) +
        '<span>Mapeamento em <strong>' + p.Percentual + '%</strong></span>' +
        '<span>· ' + plural(contarSubprocessosRecursivo(cod), 'subprocesso', 'subprocessos') + ' · ' +
        plural(totalAtivs, 'atividade', 'atividades') + ' · ' + plural(totalTarefas, 'tarefa', 'tarefas') + '</span>' +
        (p.Unidade_Organica_Responsavel ? '<span>' + esc(p.Unidade_Organica_Responsavel) + '</span>' : '') +
        (p.Processo_ECodevasf ? '<span><i class="fas fa-file-lines" aria-hidden="true"></i> ' + (p.Processo_ECodevasf_Link ? '<a href="' + esc(p.Processo_ECodevasf_Link) + '" target="_blank" rel="noopener">e-Codevasf ' + esc(p.Processo_ECodevasf) + '<span class="sr-only"> (abre em nova aba)</span></a>' : 'e-Codevasf ' + esc(p.Processo_ECodevasf)) + '</span>' : '') +
        '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do processo</h3><dl class="ficha-dl">' +
        campo('Objetivo', p.Objetivo && esc(p.Objetivo), false, 'desc') +
        campo('Descrição', p.Descricao && esc(p.Descricao), false, 'desc') +
        campo('Unidade Orgânica responsável', p.Unidade_Organica_Responsavel && siglaTag(p.Unidade_Organica_Responsavel), false, 'quem') +
        campo('Unidades Orgânicas Corresponsáveis', chips(p.Unidades_Organicas_Corresponsaveis, 'fa-people-group', true), false, 'quem') +
        campo('Ponto focal do Nugep', pontoFocalNugepHtml(p), false, 'quem') +
        campo('Gestor(a) do processo', gestorProcessoHtml(p), false, 'quem') +
        campo('Prioridade', esc(p.Prioridade || '—'), false, 'quem') +
        campo('Complexidade', esc(p.Complexidade || '—'), false, 'quem') +
        campo('Maturidade do processo', p.Maturidade && tagMaturidade(p.Maturidade), false, 'quem') +
        '<div class="span2 campo-tecnico"><dt>' + termoLink('Caminho Crítico', 'Duração estimada') + '</dt><dd>' + formatarHorasUteis(duracaoRecursivaHoras(cod)) + '</dd></div>' +
        campo('Sistemas utilizados', chips(p.Sistemas, 'fa-desktop'), false, 'tecnico') +
        campo('Fontes de dados', chips(p.Fontes_Dados, 'fa-database'), false, 'tecnico') +
        campo('Competências necessárias', chips(p.Competencias_Necessarias, 'fa-graduation-cap'), true, 'tecnico') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-right-left" aria-hidden="true"></i> ' + termoLink('SIPOC') + '</h3><div class="sipoc">' +
        '<div class="col"><h4>Fornecedores</h4><ul>' + (listar(p.Fornecedores).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Entradas</h4><ul>' + (listar(p.Entradas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col centro"><h4>Processo</h4><div style="font-weight:600">' + esc(p.Nome) + '</div></div>' +
        '<div class="col"><h4>Saídas</h4><ul>' + (listar(p.Saidas).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '<div class="col"><h4>Beneficiários</h4><ul>' + (listar(p.Beneficiarios || p.Clientes).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') || '<li class="pp-vazio">—</li>') + '</ul></div>' +
        '</div></div>' +
        '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M10)</h3>' + marcosHtml(p) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama</h3>' + diagramaHtml(p.Imagem_Bizagi, p.Nome) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Atividades ligadas direto ao processo</h3>' +
        (subs.length && !ativsDiretas.length
          ? '<p class="pp-vazio">Nenhuma atividade ligada diretamente ao processo — neste caso as atividades ficam dentro dos subprocessos listados ao lado.</p>'
          : tabelaAtividadesHtml(ativsDiretas, 'Nenhuma atividade cadastrada. Quando o processo não tem subprocessos, ligue as atividades direto a ele: coluna Vinculo_Pai da aba Atividades = ' + esc(cod) + '.', p.Unidade_Organica_Responsavel)) + '</div>' +
        secVinculos('Processo', cod) +
        '</div><aside>' +
        cardHierarquia(mp ? [{ tipo: 'mp', cat: mp._cat, codigo: mp._cod || mp.Codigo, nome: mp.Nome, href: '#/mp/' + encodeURIComponent(mp.Codigo) }] : []) +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos vinculados</h3>' +
        ((subs.length || reusados.length) ? subs.map(function (s) { return subCardHtml(s, false); }).join('') +
          reusados.map(function (s) { return subCardHtml(s, true); }).join('')
          : '<p class="pp-vazio">' + (ativsDiretas.length
          ? 'Nenhum subprocesso cadastrado — este processo se decompõe direto em atividades.'
          : 'Nenhum subprocesso cadastrado.') + '</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-forward" aria-hidden="true"></i> Próxima ação</h3>' +
        (p.Proxima_Acao
          ? '<div class="br-message info" role="status"><div class="icon"><i class="fas fa-info-circle" aria-hidden="true"></i></div>' +
            '<div class="content"><span class="message-body">' + esc(p.Proxima_Acao) + '</span></div></div>'
          : '<div class="br-message info" role="status"><div class="icon"><i class="fas fa-info-circle" aria-hidden="true"></i></div>' +
            '<div class="content"><span class="message-body">Nenhuma ação programada para este processo.</span></div></div>') +
        (p.Pendencia
          ? '<div class="br-message warning" role="alert"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div>' +
            '<div class="content"><span class="message-title">Pendência</span>' +
            '<span class="message-body">' + esc(p.Pendencia) + '</span></div></div>'
          : '') + '</div>' +
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
      var subsFilhos = IDX.subsPorPai[cod] || [];
      var subsReusados = subsReutilizadosEm(cod);
      var ativs = atividadesDe(cod);
      el.innerHTML =
        breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
          .concat(mpp ? [{ rotulo: codDisp(mpp.Codigo), href: '#/mp/' + encodeURIComponent(mpp.Codigo) }] : [])
          .concat(pp ? [{ rotulo: codDisp(pp.Codigo), href: '#/p/' + encodeURIComponent(pp.Trilha) }] : [])
          .concat(cadeiaSp.map(function (sp2) { return { rotulo: sp2.Codigo, href: '#/sp/' + encodeURIComponent(sp2.Trilha) }; }))
          .concat([{ rotulo: codDisp(s.Codigo) + ' — ' + s.Nome }])) +
        '<section class="ficha-hero nv-sp">' +
        eyebrowFicha('Subprocesso', s.Codigo) +
        '<h1>' + esc(codDisp(s.Codigo)) + ' — ' + esc(s.Nome) + '</h1>' +
        '<div class="meta"><span>' + ativs.length + ' atividades mapeadas</span>' +
        (s.Unidade_Organica_Responsavel ? '<span>· ' + esc(s.Unidade_Organica_Responsavel) + '</span>' : '') +
        '<span>· Subprocesso de ' + (cadeiaSp.length + 1) + 'º nível' + (cadeiaSp.length > 0 ? ' (aninhado)' : '') + '</span>' +
        (s.Reutilizavel === 'Sim' ? '<span class="br-tag info small"><i class="fas fa-link" aria-hidden="true"></i> Subprocesso reutilizável</span>' : '') +
        '</div></section>' +
        '<div class="ficha-grid"><div>' +
        '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha do subprocesso</h3><dl class="ficha-dl">' +
        campo('Descrição', s.Descricao && esc(s.Descricao), true, 'desc') +
        campo('Objetivo', s.Objetivo && esc(s.Objetivo), true, 'desc') +
        campo('Unidade Orgânica responsável', s.Unidade_Organica_Responsavel && siglaTag(s.Unidade_Organica_Responsavel), false, 'quem') +
        campo('Unidades Orgânicas Corresponsáveis', chips(s.Unidades_Organicas_Corresponsaveis, 'fa-people-group', true), false, 'quem') +
        campo('Entradas (insumos)', chips(s.Entradas, 'fa-arrow-right-to-bracket'), false, 'valor') +
        campo('Saídas (produtos)', chips(s.Saidas, 'fa-arrow-right-from-bracket'), false, 'valor') +
        '<div class="span2 campo-tecnico"><dt>' + termoLink('Caminho Crítico', 'Duração estimada') + '</dt><dd>' + formatarHorasUteis(duracaoRecursivaHoras(cod)) + '</dd></div>' +
        campo('Sistemas', chips(s.Sistemas, 'fa-desktop'), false, 'tecnico') +
        campo('Fontes de dados', chips(s.Fontes_Dados, 'fa-database'), false, 'tecnico') + '</dl></div>' +
        '<div class="pp-card"><h3><i class="fas fa-sitemap" aria-hidden="true"></i> Subprocessos deste subprocesso</h3>' +
        ((subsFilhos.length || subsReusados.length) ?
          tabelaGov({
            titulo: 'Subprocessos', total: subsFilhos.length + subsReusados.length,
            rotuloTotal: (subsFilhos.length + subsReusados.length) === 1 ? 'subprocesso' : 'subprocessos',
            colunas: [{ r: 'Código', curta: true, min: 132 }, { r: 'Subprocesso', principal: true }, { r: 'Ação', fixa: true, curta: true }],
            linhas: subsFilhos.map(function (sf) { return linhaSubprocesso(sf, false); }).join('') +
              subsReusados.map(function (sf) { return linhaSubprocesso(sf, true); }).join('')
          })
          : '<p class="pp-vazio">Nenhum subprocesso cadastrado dentro deste subprocesso.</p>') + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-diagram-project" aria-hidden="true"></i> Diagrama</h3>' + diagramaHtml(s.Imagem_Bizagi, s.Nome) + '</div>' +
        '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Atividades</h3>' +
        tabelaAtividadesHtml(ativs, undefined, s.Unidade_Organica_Responsavel) + '</div>' +
        secVinculos('Subprocesso', cod) +
        '</div><aside>' +
        cardHierarquia(
          (mpp ? [{ tipo: 'mp', cat: mpp._cat, codigo: mpp.Codigo, nome: mpp.Nome, href: '#/mp/' + encodeURIComponent(mpp.Codigo) }] : [])
          .concat(pp ? [{ tipo: 'p', codigo: pp.Codigo, nome: pp.Nome, href: '#/p/' + encodeURIComponent(pp.Trilha) }] : [])
          .concat(cadeiaSp.map(function (sp2) { return { tipo: 'sp', codigo: sp2.Codigo, nome: sp2.Nome, href: '#/sp/' + encodeURIComponent(sp2.Trilha) }; }))
        ) +
        // "Reutilizado em" vive aqui, logo abaixo de "Navegar para" — mesma
        // disposição da ficha do processo (Navegar para → Subprocessos
        // vinculados), padronizando as duas telas.
        (s.Reutilizavel === 'Sim' ? '<div class="pp-card"><h3><i class="fas fa-link" aria-hidden="true"></i> Reutilizado em</h3>' +
          usosDoReutilizavel(s).map(function (u) {
            return '<a class="proc-card" style="margin-bottom:var(--sp2)" href="' + u.href + '"><div class="topo"><div><span class="cod">' + esc(u.trilha) + '</span><div class="nome" style="font-size:var(--fs-sm)">' + esc(u.nome) + '</div></div></div></a>';
          }).join('') + '</div>' : '') +
        '</aside></div>';
      // clique na linha abre a atividade
      $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
        tr.addEventListener('click', function (ev) {
          if (ev.target.closest('a, input, label, button, .column-checkbox')) return;
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
    // Mesma regra da tabela: a atividade não tem unidade própria, herda a
    // do pai (subprocesso ou, na ausência dele, o processo direto).
    var unidadeAtiv = sp2 ? sp2.Unidade_Organica_Responsavel : (p2 ? p2.Unidade_Organica_Responsavel : null);
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp2 ? [{ rotulo: mp2._cod || mp2.Codigo, href: '#/mp/' + encodeURIComponent(mp2.Codigo) }] : [])
        .concat(p2 ? [{ rotulo: codDisp(p2.Codigo), href: '#/p/' + encodeURIComponent(p2.Trilha) }] : [])
        .concat(cadeiaSp2.map(function (spx) { return { rotulo: spx.Codigo, href: '#/sp/' + encodeURIComponent(spx.Trilha) }; }))
        .concat([{ rotulo: codDisp(a.Codigo) }])) +
      '<section class="ficha-hero nv-a">' +
      eyebrowFicha('Atividade', a.Codigo) +
      '<h1>' + esc(codDisp(a.Codigo)) + ' — ' + esc(a.Nome) + '</h1>' +
      '<div class="meta">' +
      (tf3.length ? '<span>· Duração estimada: ' + formatarHorasUteis(duracaoAtividadeHoras(a)) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da atividade</h3><dl class="ficha-dl">' +
      campo('Descrição', a.Descricao && esc(a.Descricao), true, 'desc') +
      campo('Tipo de atividade', a.Tipo_Atividade && tagTipoAtividade(a.Tipo_Atividade), false, 'tecnico') +
      campo('Unidade Orgânica Responsável', unidadeAtiv && siglaTag(unidadeAtiv), false, 'quem') +
      campo('Entradas (insumos)', chips(a.Entradas, 'fa-arrow-right-to-bracket'), false, 'valor') +
      campo('Saídas (produtos)', chips(a.Saidas, 'fa-arrow-right-from-bracket'), false, 'valor') +
      campo('Sistemas', chips(a.Sistemas, 'fa-desktop'), false, 'tecnico') + '</dl></div>' +
      '<div class="pp-card"><h3><i class="fas fa-list-check" aria-hidden="true"></i> Tarefas</h3>' +
      (tf3.length ? tabelaGov({
        titulo: 'Tarefas', total: tf3.length,
        rotuloTotal: tf3.length === 1 ? 'tarefa' : 'tarefas',
        colunas: [{ r: 'Código', curta: true, min: 132 }, { r: 'Tarefa', principal: true }, { r: 'Tipo' }, { r: 'Duração', curta: true }],
        linhas: tf3.map(function (t) {
          return '<tr data-link><td class="cod">' + esc(codDisp(t.Codigo)) + '</td><td><a href="#/t/' + encodeURIComponent(t.Trilha) + '"><strong>' + esc(t.Nome) + '</strong></a></td>' +
            '<td>' + esc(t.Tipo_Tarefa || '—') + '</td>' +
            '<td class="text-nowrap">' + (t.Duracao_Estimada ? formatarHorasUteis(horasTarefa(t)) : '—') + '</td></tr>';
        }).join('')
      }) : '<p class="pp-vazio">Nenhuma tarefa cadastrada para esta atividade.</p>') + '</div>' +
      secVinculos('Atividade', cod) +
      '</div><aside>' +
      cardHierarquia(
        (mp2 ? [{ tipo: 'mp', cat: mp2._cat, codigo: mp2._cod || mp2.Codigo, nome: mp2.Nome, href: '#/mp/' + encodeURIComponent(mp2.Codigo) }] : [])
        .concat(p2 ? [{ tipo: 'p', codigo: p2.Codigo, nome: p2.Nome, href: '#/p/' + encodeURIComponent(p2.Trilha) }] : [])
        .concat(cadeiaSp2.map(function (spx) { return { tipo: 'sp', codigo: spx.Codigo, nome: spx.Nome, href: '#/sp/' + encodeURIComponent(spx.Trilha) }; }))
      ) +
      '</aside></div>';
    $all('#viewDetalhe tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        // A linha é atalho para a ficha, menos onde há controle próprio:
        // link, caixa de seleção, botão de expandir.
        if (ev.target.closest('a, input, label, button, .column-checkbox')) return;
        var lk = tr.querySelector('a'); if (lk) location.hash = lk.getAttribute('href');
      });
    });
    return;
    }
    // tipo === 't' — ficha da tarefa
    var t = IDX.t[cod];
    if (!t) { el.innerHTML = naoEncontrado('Tarefa', cod); return; }
    var a3 = IDX.a[trilhaPai(t.Trilha)];
    var anc3 = ancestraisDaAtividade(a3);
    var sp3 = anc3.sp, cadeiaSp3 = anc3.cadeiaSp, p3 = anc3.p, mp3 = anc3.mp;
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Cadeia de Valor', href: '#/' }]
        .concat(mp3 ? [{ rotulo: codDisp(mp3.Codigo), href: '#/mp/' + encodeURIComponent(mp3.Codigo) }] : [])
        .concat(p3 ? [{ rotulo: codDisp(p3.Codigo), href: '#/p/' + encodeURIComponent(p3.Trilha) }] : [])
        .concat(cadeiaSp3.map(function (spx) { return { rotulo: spx.Codigo, href: '#/sp/' + encodeURIComponent(spx.Trilha) }; }))
        .concat(a3 ? [{ rotulo: codDisp(a3.Codigo), href: '#/a/' + encodeURIComponent(a3.Trilha) }] : [])
        .concat([{ rotulo: codDisp(t.Codigo) }])) +
      '<section class="ficha-hero nv-t">' +
      eyebrowFicha('Tarefa', t.Codigo) +
      '<h1>' + esc(codDisp(t.Codigo)) + ' — ' + esc(t.Nome) + '</h1>' +
      '<div class="meta">' + (t.Tipo_Tarefa ? '<span><i class="fas fa-gear" aria-hidden="true"></i> ' + esc(t.Tipo_Tarefa) + '</span>' : '') +
      (t.Duracao_Estimada ? '<span>· Duração estimada: ' + formatarHorasUteis(horasTarefa(t)) + '</span>' : '') + '</div></section>' +
      '<div class="ficha-grid"><div>' +
      '<div class="pp-card"><h3><i class="fas fa-id-card" aria-hidden="true"></i> Ficha da tarefa</h3><dl class="ficha-dl">' +
      campo('Observações', t.Observacoes && esc(t.Observacoes), true) + '</dl></div>' +
      '</div><aside>' +
      cardHierarquia(
        (mp3 ? [{ tipo: 'mp', cat: mp3._cat, codigo: mp3.Codigo, nome: mp3.Nome, href: '#/mp/' + encodeURIComponent(mp3.Codigo) }] : [])
        .concat(p3 ? [{ tipo: 'p', codigo: p3.Codigo, nome: p3.Nome, href: '#/p/' + encodeURIComponent(p3.Trilha) }] : [])
        .concat(cadeiaSp3.map(function (spx) { return { tipo: 'sp', codigo: spx.Codigo, nome: spx.Nome, href: '#/sp/' + encodeURIComponent(spx.Trilha) }; }))
        .concat(a3 ? [{ tipo: 'a', codigo: a3.Codigo, nome: a3.Nome, href: '#/a/' + encodeURIComponent(a3.Trilha) }] : [])
      ) +
      '</aside></div>';
  }
  function naoEncontrado(tipo, cod) {
    return breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: tipo + ' não encontrado' }]) +
      '<div class="template-erro">' +
      '<div class="erro-topo">' +
      '<div class="imagem-erro"><img src="img/ilustracoes/erro/error01.png" alt="" aria-hidden="true"></div>' +
      '<div class="erro-text">' +
      '<p class="erro-titulo">Não encontramos esse ' + esc(tipo.toLowerCase()) + '</p>' +
      '<p class="erro-sub">Podemos ajudá-lo a chegar ao que procura de outra forma.</p>' +
      '<p class="erro-desc">O código <strong>' + esc(cod || '—') + '</strong> não está no portfólio. ' +
      'Talvez o endereço tenha vindo com um erro de digitação, ou o item ainda não foi publicado.</p>' +
      '</div></div>' +
      '<div class="erro-busca">' + buscaCampoHtml('erroBusca', 'Aproveite para fazer uma pesquisa', 'O que você procura?', '') + '</div>' +
      '<div class="buttons-auxiliary">' +
      '<button class="br-button crumb" type="button" data-erro-acao="voltar">' +
      '<i class="fas fa-chevron-left" aria-hidden="true"></i><span>Ir para Página Anterior</span></button>' +
      '<a class="br-button crumb" href="#/"><i class="fas fa-home" aria-hidden="true"></i>' +
      '<span>Ir para Página Principal</span></a>' +
      '<button class="br-button crumb" type="button" data-erro-acao="feedback">' +
      '<i class="fas fa-comment-dots" aria-hidden="true"></i><span>Envie um Feedback</span></button>' +
      '</div></div>';
  }
  /* Ações do Template Erro. Delegação no contêiner das telas: o bloco é
     redesenhado a cada rota, e um listener por render acumularia. */
  document.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-erro-acao]');
    if (!b) return;
    if (b.getAttribute('data-erro-acao') === 'voltar') history.back();
    else { var rp = document.getElementById('btnReportError'); if (rp) rp.click(); }
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Enter') return;
    var i = ev.target.closest('#erroBusca');
    if (i && i.value.trim()) location.hash = '#/busca?q=' + encodeURIComponent(i.value.trim());
  });

  /* ── TELAS: documentos · riscos · indicadores · diário ────────────── */
  // pCod/spCod/aCod aqui são códigos BRUTOS (do filtro cumulativo de
  // Documentos) — ambíguos entre macroprocessos, resolvidos pelo primeiro
  // item com aquele código (idx.*ByCod), igual ao resto da resolução de
  // vínculo. Resolvido o objeto, o resto anda pela Trilha (única).
  function todosSubsDoProcesso(pCod) {
    var p0 = IDX.pByCod[pCod];
    var out = [], fila = p0 ? (IDX.subsPorPai[p0.Trilha] || []).slice() : [];
    while (fila.length) { var s = fila.shift(); out.push(s); (IDX.subsPorPai[s.Trilha] || []).forEach(function (c) { fila.push(c); }); }
    return out;
  }
  function todasAtividadesDoProcesso(pCod) {
    var p0 = IDX.pByCod[pCod];
    var subs = todosSubsDoProcesso(pCod), out = p0 ? (IDX.ativsPorPai[p0.Trilha] || []).slice() : [];
    subs.forEach(function (s) { out = out.concat(IDX.ativsPorPai[s.Trilha] || []); });
    return out;
  }
  function ativsDoSubBruto(spCod) { var s0 = IDX.spByCod[spCod]; return s0 ? atividadesDe(s0.Trilha) : []; }
  function tarefasDaAtivBruto(aCod) { var a0 = IDX.aByCod[aCod]; return a0 ? (IDX.tarefasPorAtiv[a0.Trilha] || []) : []; }
  // Ancestrais (macro/processo/cadeia de subprocessos/atividade/tarefa) de UM
  // código de vínculo, para os filtros cumulativos da aba Documentos.
  function ancestrosDeCodigo(nivel, codigo) {
    if (nivel === 'Macroprocesso') return { mp: codigo, p: null, sps: [], a: null, t: null };
    if (nivel === 'Processo') { var p = resolveRef('Processo', codigo); return { mp: p && p.Macroprocesso, p: p ? p.Codigo : codigo, sps: [], a: null, t: null }; }
    if (nivel === 'Subprocesso') {
      var s0 = resolveRef('Subprocesso', codigo);
      var pp = s0 && processoDoSubprocesso(s0.Trilha), cadeia = s0 ? cadeiaSubprocessos(s0.Trilha).map(function (s) { return s.Codigo; }) : [];
      return { mp: pp && pp.Macroprocesso, p: pp && pp.Codigo, sps: cadeia, a: null, t: null };
    }
    if (nivel === 'Atividade') {
      var a = resolveRef('Atividade', codigo), an = a && ancestraisDaAtividade(a);
      var sps2 = an ? (an.cadeiaSp || []).map(function (s) { return s.Codigo; }).concat(an.sp ? [an.sp.Codigo] : []) : [];
      return { mp: an && an.mp && an.mp.Codigo, p: an && an.p && an.p.Codigo, sps: sps2, a: a ? a.Codigo : codigo, t: null };
    }
    if (nivel === 'Tarefa') {
      var t = resolveRef('Tarefa', codigo), atv = t && resolveRef('Atividade', t.Atividade), an2 = atv && ancestraisDaAtividade(atv);
      var sps3 = an2 ? (an2.cadeiaSp || []).map(function (s) { return s.Codigo; }).concat(an2.sp ? [an2.sp.Codigo] : []) : [];
      return { mp: an2 && an2.mp && an2.mp.Codigo, p: an2 && an2.p && an2.p.Codigo, sps: sps3, a: atv ? atv.Codigo : (t && t.Atividade), t: t ? t.Codigo : codigo };
    }
    return { mp: null, p: null, sps: [], a: null, t: null };
  }
  function docAncestros(x) {
    var niv = x.Vinculo_Nivel, codigos = String(x.Vinculo_Codigo || '').split(';').map(function (s) { return s.trim(); }).filter(Boolean);
    var out = { mp: {}, p: {}, sp: {}, a: {}, t: {} };
    codigos.forEach(function (c) {
      var an = ancestrosDeCodigo(niv, c);
      if (an.mp) out.mp[an.mp] = 1;
      if (an.p) out.p[an.p] = 1;
      (an.sps || []).forEach(function (s) { out.sp[s] = 1; });
      if (an.a) out.a[an.a] = 1;
      if (an.t) out.t[an.t] = 1;
    });
    return out;
  }
  var filtroDoc = { tipo: '', situacao: '', fmp: '', fp: '', fsp: '', fa: '', ft: '', q: '', busca: '', ordem: '', dir: '', dens: 'medium', sel: {} };
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
    { k: 'ID', r: 'ID', min: 108 }, { k: 'Titulo', r: 'Documento', min: 420 },
    { k: 'Vinculo_Codigo', r: 'Vinculado a', min: 200 }, { k: 'Data', r: 'Data' },
    { k: 'Situacao', r: 'Situação' }
  ];
  function renderDocumentos() {
    var el = $('#viewDocumentos');
    var tipos = {};
    DADOS.docs.forEach(function (x) { if (x.Tipo_Documento) tipos[x.Tipo_Documento] = 1; });
    var situacoes = {};
    DADOS.docs.forEach(function (x) { if (x.Situacao) situacoes[x.Situacao] = 1; });
    var docsAnc = {}; DADOS.docs.forEach(function (x) { docsAnc[x.ID] = docAncestros(x); });
    // Opções cumulativas: cada nível só lista o que existe dentro do nível pai já escolhido.
    var opMp = DADOS.macros;
    var opP = filtroDoc.fmp ? (IDX.procsPorMacro[filtroDoc.fmp] || []) : DADOS.procs;
    var opSp = filtroDoc.fp ? todosSubsDoProcesso(filtroDoc.fp) : (filtroDoc.fmp ? DADOS.subs.filter(function (s) { var pp = processoDoSubprocesso(s.Trilha); return pp && pp.Macroprocesso === filtroDoc.fmp; }) : DADOS.subs);
    var opA = filtroDoc.fsp ? ativsDoSubBruto(filtroDoc.fsp) : filtroDoc.fp ? todasAtividadesDoProcesso(filtroDoc.fp) : DADOS.ativs;
    var opT = filtroDoc.fa ? tarefasDaAtivBruto(filtroDoc.fa) : DADOS.tarefas;
    var bq = filtroDoc.busca.toLowerCase();
    var lista = DADOS.docs.filter(function (x) {
      if (filtroDoc.tipo && x.Tipo_Documento !== filtroDoc.tipo) return false;
      if (filtroDoc.situacao && x.Situacao !== filtroDoc.situacao) return false;
      var anc = docsAnc[x.ID];
      if (filtroDoc.fmp && !anc.mp[filtroDoc.fmp]) return false;
      if (filtroDoc.fp && !anc.p[filtroDoc.fp]) return false;
      if (filtroDoc.fsp && !anc.sp[filtroDoc.fsp]) return false;
      if (filtroDoc.fa && !anc.a[filtroDoc.fa]) return false;
      if (filtroDoc.ft && !anc.t[filtroDoc.ft]) return false;
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
      '<div class="pp-sec-h" style="margin-top:0"><h1>Repositório de documentos</h1><div class="linha" aria-hidden="true"></div></div>' +
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros do repositório de documentos">' +
      '<div class="filtros-campos">' +
      selectHtml({ chave: 'tipoDoc', id: 'fTipoDoc', rotulo: 'Tipo de documento',
        placeholder: 'Todos os tipos', selecionados: filtroDoc.tipo ? [filtroDoc.tipo] : [],
        opcoes: Object.keys(tipos).sort().map(function (t) { return { v: t, r: t }; }) }) +
      selectHtml({ chave: 'situacaoDoc', id: 'fSituacaoDoc', rotulo: 'Situação',
        placeholder: 'Todas as situações', selecionados: filtroDoc.situacao ? [filtroDoc.situacao] : [],
        opcoes: Object.keys(situacoes).sort().map(function (t) { return { v: t, r: t }; }) }) +
      selectHtml({ chave: 'fmpDoc', id: 'fFmpDoc', rotulo: 'Macroprocesso',
        placeholder: 'Todos os macroprocessos', selecionados: filtroDoc.fmp ? [filtroDoc.fmp] : [],
        opcoes: opMp.map(function (m) { return { v: m.Codigo, r: codDisp(m.Codigo) + ' — ' + m.Nome }; }) }) +
      selectHtml({ chave: 'fpDoc', id: 'fFpDoc', rotulo: 'Processo',
        placeholder: 'Todos os processos', selecionados: filtroDoc.fp ? [filtroDoc.fp] : [],
        opcoes: opP.map(function (p) { return { v: p.Codigo, r: codDisp(p.Codigo) + ' — ' + p.Nome }; }) }) +
      selectHtml({ chave: 'fspDoc', id: 'fFspDoc', rotulo: 'Subprocesso',
        placeholder: 'Todos os subprocessos', selecionados: filtroDoc.fsp ? [filtroDoc.fsp] : [],
        opcoes: opSp.map(function (s) { return { v: s.Codigo, r: codDisp(s.Codigo) + ' — ' + s.Nome }; }) }) +
      selectHtml({ chave: 'faDoc', id: 'fFaDoc', rotulo: 'Atividade',
        placeholder: 'Todas as atividades', selecionados: filtroDoc.fa ? [filtroDoc.fa] : [],
        opcoes: opA.map(function (a) { return { v: a.Codigo, r: codDisp(a.Codigo) + ' — ' + a.Nome }; }) }) +
      selectHtml({ chave: 'ftDoc', id: 'fFtDoc', rotulo: 'Tarefa',
        placeholder: 'Todas as tarefas', selecionados: filtroDoc.ft ? [filtroDoc.ft] : [],
        opcoes: opT.map(function (t) { return { v: t.Codigo, r: codDisp(t.Codigo) + ' — ' + t.Nome }; }) }) +
      '</div>' +
      (filtroDoc.tipo || filtroDoc.situacao || filtroDoc.fmp || filtroDoc.fp || filtroDoc.fsp || filtroDoc.fa || filtroDoc.ft || filtroDoc.busca
        ? '<div class="filtros-rodape"><div class="filtros-ativos">' +
          (filtroDoc.tipo ? '<span class="br-tag interaction" id="fdoc-tipo"><span id="fdoc-tipo-r">Tipo: ' + esc(filtroDoc.tipo) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="tipo"' +
            ' aria-label="Remover o filtro de tipo" aria-describedby="fdoc-tipo-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.situacao ? '<span class="br-tag interaction" id="fdoc-situacao"><span id="fdoc-situacao-r">Situação: ' + esc(filtroDoc.situacao) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="situacao"' +
            ' aria-label="Remover o filtro de situação" aria-describedby="fdoc-situacao-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.fmp ? '<span class="br-tag interaction" id="fdoc-fmp"><span id="fdoc-fmp-r">Macroprocesso: ' + esc(codDisp(filtroDoc.fmp)) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="fmp" aria-label="Remover o filtro de macroprocesso" aria-describedby="fdoc-fmp-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.fp ? '<span class="br-tag interaction" id="fdoc-fp"><span id="fdoc-fp-r">Processo: ' + esc(codDisp(filtroDoc.fp)) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="fp" aria-label="Remover o filtro de processo" aria-describedby="fdoc-fp-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.fsp ? '<span class="br-tag interaction" id="fdoc-fsp"><span id="fdoc-fsp-r">Subprocesso: ' + esc(codDisp(filtroDoc.fsp)) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="fsp" aria-label="Remover o filtro de subprocesso" aria-describedby="fdoc-fsp-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.fa ? '<span class="br-tag interaction" id="fdoc-fa"><span id="fdoc-fa-r">Atividade: ' + esc(codDisp(filtroDoc.fa)) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="fa" aria-label="Remover o filtro de atividade" aria-describedby="fdoc-fa-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.ft ? '<span class="br-tag interaction" id="fdoc-ft"><span id="fdoc-ft-r">Tarefa: ' + esc(codDisp(filtroDoc.ft)) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="ft" aria-label="Remover o filtro de tarefa" aria-describedby="fdoc-ft-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          (filtroDoc.busca ? '<span class="br-tag interaction" id="fdoc-busca"><span id="fdoc-busca-r">Pesquisa: ' + esc(filtroDoc.busca) + '</span>' +
            '<button class="br-button circle small" type="button" data-doc-limpar="busca"' +
            ' aria-label="Remover o termo pesquisado" aria-describedby="fdoc-busca-r">' +
            '<i class="fas fa-times" aria-hidden="true"></i></button></span>' : '') +
          '</div><button class="br-button secondary small" type="button" data-doc-limpar="tudo">' +
          '<i class="fas fa-rotate-left" aria-hidden="true"></i> Limpar Filtros</button></div>'
        : '') +
      '</section>' +
      '<div class="br-table ' + esc(filtroDoc.dens) + '" id="tabelaDocs" data-search="data-search" data-selection="data-selection" data-collapse="data-collapse">' +
      /* Barra de Título (item 1): título + ações utilitárias. Acima de 4
         ações a spec pede menu flutuante — densidade vai no ellipsis-v. */
      '<div class="table-header">' +
      '<div class="top-bar">' +
      '<div class="table-title">Documentos publicados</div>' +
      '<div class="actions-trigger text-nowrap">' +
      '<span class="tooltip-wrap">' +
      '<button class="br-button circle small" type="button" id="docsDensBtn" data-tooltip-trigger data-toggle="dropdown" data-target="docsDensMenu" aria-label="Definir densidade da tabela" aria-haspopup="true" aria-expanded="false"><i class="fas fa-ellipsis-v" aria-hidden="true"></i></button>' +
      '<span class="br-tooltip small" role="tooltip" info place="bottom">' +
      '<span class="text">Densidade da tabela</span>' +
      '<span class="subtext">Compacta, regular ou espaçada</span></span>' +
      '</span>' +
      '<div class="br-list dd-target" id="docsDensMenu" role="menu" aria-labelledby="docsDensBtn" hidden>' +
      [['small', 'Compacta'], ['medium', 'Regular'], ['large', 'Espaçada']].map(function (o, i) {
        return (i ? '<span class="br-divider" role="presentation"></span>' : '') +
          '<button class="br-item" type="button" role="menuitem" data-density="' + o[0] + '"' +
          (filtroDoc.dens === o[0] ? ' aria-current="true"' : '') + '>' + o[1] + '</button>';
      }).join('') + '</div></div>' +
      '<div class="search-trigger"><button class="br-button circle small" type="button" id="docsSearchBtn" data-toggle="search" aria-label="Abrir busca" aria-controls="docsSearchInput" aria-expanded="false"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '</div>' +
      /* Busca (Comportamento 6): cobre a barra de título enquanto ativa. */
      '<div class="search-bar' + (filtroDoc.busca ? ' show' : '') + '">' +
      '<div class="br-input"><label for="docsSearchInput">Pesquisar na tabela</label>' +
      '<input id="docsSearchInput" type="search" placeholder="Pesquisar na tabela" value="' + esc(filtroDoc.busca) + '">' +
      '<button class="br-button circle small" type="button" aria-label="Pesquisar"><i class="fas fa-search" aria-hidden="true"></i></button></div>' +
      '<button class="br-button circle small" type="button" data-dismiss="search" aria-label="Fechar pesquisa"><i class="fas fa-times" aria-hidden="true"></i></button>' +
      '</div>' +
      /* Barra Contextual (item 2): surge sob a barra de título ao
         selecionar linhas, com contagem e ações contextuais. */
      '<div class="selected-bar' + (nSel ? ' show' : '') + '">' +
      '<div class="info"><span class="count">' + nSel + '</span><span class="text">' + (nSel === 1 ? 'item selecionado' : 'itens selecionados') + '</span></div>' +
      '<span class="tooltip-wrap">' +
      '<button class="br-button circle small" type="button" id="docsSelExport" aria-label="Exportar seleção em CSV" data-tooltip-trigger><i class="fas fa-file-csv" aria-hidden="true"></i></button>' +
      '<span class="br-tooltip small" role="tooltip" info place="bottom"><span class="subtext">Exportar seleção em CSV</span></span>' +
      '</span>' +
      '<span class="tooltip-wrap">' +
      '<button class="br-button circle small" type="button" id="docsSelClear" aria-label="Limpar seleção" data-tooltip-trigger><i class="fas fa-times" aria-hidden="true"></i></button>' +
      '<span class="br-tooltip small" role="tooltip" info place="bottom"><span class="subtext">Limpar seleção</span></span>' +
      '</span>' +
      '</div></div>' +
      (pagina.length ? '<div class="responsive"><table><caption>Documentos publicados</caption><thead><tr>' +
      '<td class="column-collapse" aria-hidden="true"></td>' +
      '<th class="column-checkbox" scope="col"><div class="br-checkbox hidden-label">' +
      '<input id="docsCheckAll" name="docsCheckAll" type="checkbox" aria-label="Selecionar tudo"' + (todosSel ? ' checked' : '') + '>' +
      '<label for="docsCheckAll">Selecionar todas as linhas</label></div></th>' +
      COLS_DOC.map(function (c) {
        var dir = filtroDoc.ordem === c.k ? filtroDoc.dir : '';
        var ic = dir === 'asc' ? 'fa-sort-up' : dir === 'desc' ? 'fa-sort-down' : 'fa-sort';
        // min: largura mínima da coluna (esta tabela não tem colgroup, o
        // th é quem carrega a medida).
        return '<th scope="col"' + (dir ? ' aria-sort="' + (dir === 'asc' ? 'ascending' : 'descending') + '"' : '') +
          (c.min ? ' style="min-width:' + c.min + 'px"' : '') + '>' +
          '<button type="button" class="sort-btn" data-sort="' + c.k + '" aria-label="Ordenar por ' + esc(c.r) + '">' +
          esc(c.r) + '<i class="fas ' + ic + '" aria-hidden="true"></i></button></th>';
      }).join('') + '</tr></thead><tbody>' +
      pagina.map(function (x, i) {
        var tit = x.Link ? '<a href="' + esc(x.Link) + '" target="_blank" rel="noopener">' + esc(x.Titulo) + '<span class="sr-only"> (abre em nova aba)</span></a>' : esc(x.Titulo);
        var cid = 'doc-col-' + i, chk = 'doc-chk-' + i, marcada = !!filtroDoc.sel[x.ID];
        return '<tr' + (marcada ? ' class="is-selected"' : '') + '>' +
          '<td class="column-collapse"><button class="br-button circle small" type="button" data-toggle="collapse" data-target="' + cid + '" aria-expanded="false" aria-controls="' + cid + '" aria-label="Expandir ou retrair ' + esc(x.Titulo) + '"><i class="fas fa-chevron-down" aria-hidden="true"></i></button></td>' +
          '<td><div class="br-checkbox hidden-label"><input id="' + chk + '" name="' + chk + '" type="checkbox" data-doc="' + esc(x.ID) + '" aria-label="Selecionar ' + esc(x.Titulo) + '"' + (marcada ? ' checked' : '') + '><label for="' + chk + '">Selecionar linha</label></div></td>' +
          '<td class="cod" data-th="ID">' + esc(x.ID) + '</td>' +
          '<td data-th="Documento"><strong>' + tit + '</strong></td>' +
          '<td data-th="Vinculado a">' + linkVinculos(x.Vinculo_Nivel, x.Vinculo_Codigo) + '</td>' +
          '<td data-th="Data">' + fmtData(x.Data) + '</td>' +
          '<td data-th="Situação">' + esc(x.Situacao || '—') + (revisaoVencida(x) ? ' <span class="br-tag revisao-vencida"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> Revisão vencida</span>' : '') + '</td></tr>' +
          '<tr class="collapse"><td id="' + cid + '" colspan="7" hidden>' +
          '<div class="br-list" role="list">' +
          '<div class="br-item" role="listitem"><strong>Tipo:</strong> ' + esc(x.Tipo_Documento || '—') + '</div>' +
          '<div class="br-item" role="listitem"><strong>Versão:</strong> ' + esc(x.Versao || '—') + '</div>' +
          (x.Ato_Aprovacao ? '<div class="br-item" role="listitem"><strong>Ato de aprovação:</strong> ' + esc(x.Ato_Aprovacao) + '</div>' : '') +
          '</div></td></tr>';
      }).join('') +
      '</tbody></table></div>' +
      '<div class="table-footer">' + paginacaoHtml('docs', lista.length, 'documentos') + '</div>'
        : vazio('Nenhum documento com esses filtros',
          'Revise o tipo selecionado ou o termo buscado para ver os documentos disponíveis.',
          { img: 'empty-space/empty-space-03.png' })) + '</div>';
    ligarPaginacao(el, renderDocumentos);
    el.querySelectorAll('[data-doc-limpar]').forEach(function (b) {
      b.onclick = function () {
        var q = b.getAttribute('data-doc-limpar');
        if (q === 'tipo' || q === 'tudo') filtroDoc.tipo = '';
        if (q === 'situacao' || q === 'tudo') filtroDoc.situacao = '';
        if (q === 'fmp' || q === 'tudo') { filtroDoc.fmp = ''; filtroDoc.fp = ''; filtroDoc.fsp = ''; filtroDoc.fa = ''; filtroDoc.ft = ''; }
        if (q === 'fp' || q === 'tudo') { filtroDoc.fp = ''; filtroDoc.fsp = ''; filtroDoc.fa = ''; filtroDoc.ft = ''; }
        if (q === 'fsp' || q === 'tudo') { filtroDoc.fsp = ''; filtroDoc.fa = ''; filtroDoc.ft = ''; }
        if (q === 'fa' || q === 'tudo') { filtroDoc.fa = ''; filtroDoc.ft = ''; }
        if (q === 'ft' || q === 'tudo') filtroDoc.ft = '';
        if (q === 'busca' || q === 'tudo') filtroDoc.busca = '';
        PAG.docs.pag = 1;
        renderDocumentos();
      };
    });
    window.BRSelectInit(el, function (chave, valores) {
      var v = valores[0] || '';
      if (chave === 'tipoDoc') filtroDoc.tipo = v;
      else if (chave === 'situacaoDoc') filtroDoc.situacao = v;
      else if (chave === 'fmpDoc') { filtroDoc.fmp = v; filtroDoc.fp = ''; filtroDoc.fsp = ''; filtroDoc.fa = ''; filtroDoc.ft = ''; }
      else if (chave === 'fpDoc') { filtroDoc.fp = v; filtroDoc.fsp = ''; filtroDoc.fa = ''; filtroDoc.ft = ''; }
      else if (chave === 'fspDoc') { filtroDoc.fsp = v; filtroDoc.fa = ''; filtroDoc.ft = ''; }
      else if (chave === 'faDoc') { filtroDoc.fa = v; filtroDoc.ft = ''; }
      else filtroDoc.ft = v;
      PAG.docs.pag = 1;
      renderDocumentos();
    });

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
          return COLS_DOC.map(function (c) {
            var v = x[c.k];
            if (c.k === 'Vinculo_Codigo') v = String(v || '').split(';')
              .map(function (cd) { return codDisp(cd.trim()); }).join('; ');
            return String(v == null ? '' : v);
          });
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
            return '<span class="tooltip-wrap"><button type="button" class="risco-pin" data-tooltip-trigger' +
              ' data-alvo="risco-' + esc(r.ID) + '" aria-label="' + esc(r.ID + ' — ' + r.Descricao_Risco) + '">' +
              esc(r.ID.replace(/^RIS-/i, '')) + '</button>' +
              '<span class="br-tooltip small" role="tooltip" info place="top"><span class="text">' + esc(r.ID) +
              '</span><span class="subtext">' + esc(r.Descricao_Risco) + '</span></span></span>';
          }).join('');
        celulas += '<div class="cel ' + cls + '">' + pins + '</div>';
      }
    }
    celulas += '<div class="cab"></div>';
    for (var pr = 1; pr <= 5; pr++) celulas += '<div class="cab">' + pr + '</div>';
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>Radar de riscos</h1><div class="linha" aria-hidden="true"></div></div>' +
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
    // Agrupado por Categoria (SLA/ROI/Processo/Produto/Projeto) — mapeia
    // quase 1:1 às perspectivas do Balanced Scorecard (CBOK 7.3.3/10.1.9):
    // SLA≈Cliente, Financeiro (ROI)≈Financeira, Processo≈Processos Internos,
    // Produto/Projeto≈Aprendizado e Crescimento.
    var ORDEM_CAT = ['Nível de Serviço (SLA)', 'Financeiro (ROI)', 'Processo', 'Produto', 'Projeto'];
    var porCategoria = {};
    DADOS.metricas.forEach(function (x) { var c = x.Categoria || 'Sem categoria'; (porCategoria[c] = porCategoria[c] || []).push(x); });
    var categorias = ORDEM_CAT.filter(function (c) { return porCategoria[c]; })
      .concat(Object.keys(porCategoria).filter(function (c) { return ORDEM_CAT.indexOf(c) < 0; }));
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>Indicadores de desempenho</h1><div class="linha" aria-hidden="true"></div></div>' +
      categorias.map(function (c) {
        return '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> ' + esc(c) + '</h3>' + tabelaIndsHtml(porCategoria[c], true) + '</div>';
      }).join('');
  }
  /* ── TELA: metodologia ────────────────────────────────────────────── */
  /* ── GRÁFICOS (SVG puro, sem dependências; cores do DS gov.br) ────── */
  // Paleta categórica dos gráficos: as sete matizes do sistema de camadas
  // (azul, verde, laranja, turquesa, índigo, bronze, grafite) mais o
  // vermelho de erro — todas degraus oficiais da paleta gov.br, todas com
  // 4,5:1 ou mais sobre o card branco. Matizes diferentes entre si, para
  // que séries vizinhas nunca se confundam.
  var PAL = ['#005ca9', '#007d4e', '#222b54', '#0081a1', '#74c9ea', '#89bd2b', '#555555', '#b50909'];
  function svgWrap(titulo, conteudo, vb, altura, legenda) {
    return '<figure class="graf"><figcaption>' + esc(titulo) +
      (legenda ? ' ' + dica(legenda) : '') + '</figcaption>' +
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
        '<rect x="' + lw + '" y="' + y + '" width="' + w.toFixed(1) + '" height="15" rx="3" fill="' + (d.cor || '#005ca9') + '"' + navAttr + '>' +
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
      return '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="3.5" fill="#005ca9"><title>' + esc(pontos[i].rotulo) + ': ' + pontos[i].valor + '</title></circle>' +
        '<text x="' + p[0].toFixed(1) + '" y="' + (h - 6) + '" font-size="9" fill="#636363" text-anchor="middle">' + esc(pontos[i].rotulo) + '</text>';
    }).join('');
    return svgWrap(titulo, '<path d="' + area + '" fill="#005ca9" opacity="0.12"></path>' +
      '<path d="' + linha + '" fill="none" stroke="#005ca9" stroke-width="2.5"></path>' + marcas, '0 0 ' + w + ' ' + h, null, legenda);
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
    var corAtual = '#005ca9';
    (faixas || []).forEach(function (f) { if (valor <= f.ate) corAtual = corAtual === '#005ca9' ? f.cor : corAtual; });
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
      return '<circle cx="' + px(p.x).toFixed(1) + '" cy="' + py(p.y).toFixed(1) + '" r="' + (p.r || 8) + '" fill="' + (p.cor || '#005ca9') + '" fill-opacity=".78" stroke="#fff" stroke-width="1.5"' + navAttr + '>' +
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
    // Taxa de atualização (CBOK 4.2.11, uma das 3 métricas de repositório):
    // % de processos com Ultima_Atualizacao nos últimos 90 dias.
    var corte90 = new Date(hoje + 'T00:00:00'); corte90.setDate(corte90.getDate() - 90);
    var corte90ISO = corte90.toISOString().slice(0, 10);
    var atualizRecente = procs.filter(function (p) { return p.Ultima_Atualizacao && p.Ultima_Atualizacao >= corte90ISO; }).length;
    var txAtualizPct = procs.length ? Math.round((atualizRecente / procs.length) * 100) : 0;
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

    // Cores da marca Codevasf (mesmos tons de --cv-blue/--cv-green/--cv-navy
    // em css/govbr-ds.css): azul = gerencial, verde = finalístico, azul-marinho = suporte.
    var CAT_COR = { gerencial: '#005ca9', finalistico: '#007d4e', suporte: '#222b54' };
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>Dashboard gerencial</h1><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="kpi-grid" style="margin-top:0">' +
      '<div class="kpi"><span class="num">' + cobertura + '%</span><span class="lbl">Processos publicados' + dica('Processos com todos os marcos do mapeamento (M1–M10) concluídos, em relação ao total do portfólio.') + '</span><span class="sub">' + concl + ' de ' + procs.length + ' processos</span></div>' +
      '<div class="kpi"><span class="num">' + media + '%</span><span class="lbl">Avanço médio' + dica('Percentual médio de execução do mapeamento entre todos os processos do portfólio.') + '</span><span class="sub">' + andam + ' em andamento</span></div>' +
      '<div class="kpi"><span class="num">' + procs.length + '</span><span class="lbl">Processos' + dica('Total de processos mapeados no portfólio.') + '</span><span class="sub">mapeados no portfólio</span></div>' +
      '<div class="kpi"><span class="num">' + DADOS.subs.length + '</span><span class="lbl">Subprocessos' + dica('Total de subprocessos mapeados no portfólio (inclusive os aninhados, subprocesso dentro de subprocesso).') + '</span><span class="sub">mapeados no portfólio</span></div>' +
      '<div class="kpi"><span class="num">' + DADOS.ativs.length + '</span><span class="lbl">Atividades' + dica('Total de atividades mapeadas no portfólio.') + '</span><span class="sub">mapeadas no portfólio</span></div>' +
      '<div class="kpi"><span class="num">' + DADOS.tarefas.length + '</span><span class="lbl">Tarefas' + dica('Total de tarefas mapeadas no portfólio (menor unidade de trabalho, CBOK 4.0).') + '</span><span class="sub">mapeadas no portfólio</span></div>' +
      '<div class="kpi ' + (criticos.length ? 'erro' : 'ok') + '"><span class="num">' + criticos.length + '</span><span class="lbl">Riscos críticos' + dica('Riscos classificados como Alto ou Extremo, ainda não encerrados.') + '</span><span class="sub">' + riscosAb.length + ' riscos abertos no total</span></div>' +
      '<div class="kpi"><span class="num">' + VISITAS_NAVEGADOR + '</span><span class="lbl">Visitas ao repositório' + dica('Utilização é a métrica chave de um bom repositório de processos (CBOK). Contador local deste navegador — o painel é um site estático sem login, não soma visitas de outras pessoas nem de outros dispositivos.') + '</span><span class="sub">neste navegador</span></div>' +
      '<div class="kpi"><span class="num">' + txAtualizPct + '%</span><span class="lbl">Taxa de atualização' + dica('Uma das três métricas de repositório do CBOK (4.2.11): percentual de processos atualizados nos últimos 90 dias.') + '</span><span class="sub">processos atualizados nos últimos 90 dias</span></div>' +
      '</div>' +
      '<div class="graf-grid">' +
      grafGauge('Avanço médio geral do portfólio', media, 100, [
        { ate: 40, cor: '#b50909' }, { ate: 70, cor: '#947100' }, { ate: 100, cor: '#007d4e' }],
        'Percentual médio de avanço do mapeamento entre todos os processos do portfólio. As faixas de cor indicam o estágio geral: vermelho abaixo de 40%, âmbar entre 40% e 70%, verde a partir de 70%.') +
      grafDonut('Situação do mapeamento', [
        { rotulo: 'Concluído', valor: concl, cor: '#007d4e' },
        { rotulo: 'Em andamento', valor: andam, cor: '#947100' },
        { rotulo: 'Não iniciado', valor: porStatus['Não iniciado'] || 0, cor: '#757575' },
        { rotulo: 'Suspenso', valor: porStatus['Suspenso'] || 0, cor: '#b50909' }],
        'Quantidade de processos em cada status de mapeamento. O tamanho de cada fatia é proporcional ao número de processos; passe o mouse sobre uma fatia para ver o total e o percentual.') +
      grafDonut('Processos por tipo (CBOK)', ['gerencial', 'finalistico', 'suporte'].map(function (c) {
        return { rotulo: c === 'finalistico' ? 'Finalístico' : c === 'gerencial' ? 'Gerencial' : 'Suporte',
          valor: procs.filter(function (p) { var m = IDX.mp[p.Macroprocesso]; return m && m._cat === c; }).length, cor: CAT_COR[c] };
      }), 'Distribuição dos processos mapeados entre os três tipos de macroprocesso do CBOK 4.0: gerencial, finalístico e de suporte. Referência do CBOK para comparação (não é meta obrigatória): cerca de 20% primários (finalísticos), 70% de suporte e 10% de gestão.') +
      grafBarras('Avanço por macroprocesso (%) — clique para abrir', DADOS.macros.map(function (m) {
        var ps = IDX.procsPorMacro[m.Codigo] || [];
        return { rotulo: (m._cod || m.Codigo) + ' ' + m.Nome, cor: CAT_COR[m._cat] || '#005ca9', href: '#/mp/' + encodeURIComponent(m.Codigo),
          valor: ps.length ? Math.round(ps.reduce(function (s, p) { return s + p.Percentual; }, 0) / ps.length) : 0 };
      }), '%', 100, 'Percentual médio de avanço dos processos de cada macroprocesso. A linha tracejada marca a meta de 100%; clique em uma barra para abrir a ficha do macroprocesso.') +
      grafBubble('Priorização: avanço × riscos abertos por processo — clique para abrir', procs.map(function (p) {
        var rp = vinculados('riscos', 'Processo', p.Trilha).filter(function (r) { return !/encerrad/i.test(String(r.Status || '')); });
        var ativs = contarAtividadesRecursivo(p.Trilha);
        return { x: p.Percentual, y: rp.length, r: Math.max(6, Math.min(22, 6 + ativs * 1.5)),
          rotulo: codDisp(p.Codigo) + ' — ' + p.Nome, cor: p.Percentual < 40 && rp.length >= 1 ? '#b50909' : (CAT_COR[(IDX.mp[p.Macroprocesso] || {})._cat] || '#005ca9'),
          href: '#/p/' + encodeURIComponent(p.Trilha) };
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
      '<div class="pp-card"><h3><i class="fas fa-exclamation-triangle" aria-hidden="true"></i> Processos com prazo vencido</h3>' +
      (atrasados.length ? tabelaGov({
        titulo: 'Processos com prazo vencido', total: atrasados.length,
        rotuloTotal: atrasados.length === 1 ? 'processo' : 'processos',
        colunas: [{ r: 'Código', curta: true }, { r: 'Processo', principal: true }, { r: 'Responsável' }, { r: 'Prazo', curta: true }, { r: 'Avanço', curta: true }],
        linhas: atrasados.map(function (p) {
          return '<tr data-link><td class="cod">' + esc(codDisp(p.Codigo)) + '</td><td><a href="#/p/' + encodeURIComponent(p.Trilha) + '"><strong>' + esc(p.Nome) + '</strong></a></td>' +
            '<td>' + esc(p.Ponto_Focal_Nugep || '—') + '</td><td class="text-nowrap">' + fmtData(p.Prazo_Previsto) + '</td>' +
            '<td style="min-width:120px">' + barraPct(p.Percentual) + '</td></tr>';
        }).join('')
      }) : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Nenhum processo com prazo vencido.</span></div></div>') + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-shield-halved" aria-hidden="true"></i> Riscos críticos abertos</h3>' +
      (criticos.length ? tabelaRiscosHtml(criticos, true) : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Nenhum risco crítico em aberto.</span></div></div>') + '</div>' +
      '<div class="pp-card"><h3><i class="fas fa-chart-line" aria-hidden="true"></i> Indicadores fora da meta</h3>' +
      (function () {
        var fora = DADOS.metricas.filter(function (x) { return x._sit === 'Meta não atingida'; });
        return fora.length ? tabelaIndsHtml(fora, true) : '<div class="br-message success" role="status"><div class="icon"><i class="fas fa-check-circle" aria-hidden="true"></i></div><div class="content"><span class="message-body">Todos os indicadores medidos estão na meta.</span></div></div>';
      })() + '</div></section>';
    $all('#viewDashboard tr[data-link]').forEach(function (tr) {
      tr.addEventListener('click', function (ev) {
        // A linha é atalho para a ficha, menos onde há controle próprio:
        // link, caixa de seleção, botão de expandir.
        if (ev.target.closest('a, input, label, button, .column-checkbox')) return;
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

  /* Rodapé de filtros ativos — a tela sempre mostra o que está aplicado e
     como desfazer. Sem isso, um filtro escolhido continuava valendo mesmo
     depois de o usuário sair da aba e voltar, sem nada em tela que o
     explicasse. Os pares são [chave, rótulo] dos recortes ativos. */
  function rodapeFiltrosHtml(marca, pares) {
    var ativos = pares.filter(function (p) { return p[1]; });
    if (!ativos.length) return "";
    return '<div class="filtros-rodape"><div class="filtros-ativos">' +
      ativos.map(function (p, i) {
        var rid = marca + "-f" + i;
        return '<span class="br-tag interaction"><span id="' + rid + '">' + esc(p[1]) + '</span>' +
          '<button class="br-button circle small" type="button" data-limpar-filtro="' + esc(p[0]) + '"' +
          ' data-limpar-marca="' + marca + '" aria-label="Remover este filtro" aria-describedby="' + rid + '">' +
          '<i class="fas fa-times" aria-hidden="true"></i></button></span>';
      }).join("") +
      '</div><button class="br-button secondary small" type="button" data-limpar-filtro="tudo"' +
      ' data-limpar-marca="' + marca + '"><i class="fas fa-rotate-left" aria-hidden="true"></i> Limpar Filtros</button></div>';
  }
  function ligarRodapeFiltros(el, marca, aplicar) {
    $all('[data-limpar-marca="' + marca + '"]', el).forEach(function (b) {
      b.onclick = function () { aplicar(b.getAttribute("data-limpar-filtro")); };
    });
  }

  /* ── TELA: repositório (jornada + materiais + metodologia) ────────── */
  var filtroRepo = { cat: '', fase: '', q: '' };
  function cardRepo(it) {
    var interno = !/^https?:/i.test(String(it.Link || ''));
    var icone = { 'Documento oficial': 'fa-scale-balanced', 'Template': 'fa-file-lines',
      'Instrumento': 'fa-toolbox', 'Ferramenta': 'fa-screwdriver-wrench', 'Referência': 'fa-book' }[it.Categoria] || 'fa-file';
    return '<article class="repo-card"><div class="repo-topo">' +
      '<span class="repo-ico"><i class="fas ' + icone + '" aria-hidden="true"></i></span>' +
      '<div><span class="repo-cat">' + esc(it.Categoria || '') + (it.Fase_Ciclo ? ' · ' + esc(it.Fase_Ciclo) : '') + '</span></div></div>' +
      '<h4>' + esc(it.Titulo) + '</h4><p>' + esc(it.Descricao || '') + '</p>' +
      '<div class="repo-rodape"><span class="repo-fonte">Fonte: ' + esc(it.Fonte || '—') + '</span>' +
      (it.Link ? '<a class="br-button secondary small" href="' + esc(it.Link) + '"' +
        (interno ? ' download' : ' target="_blank" rel="noopener"') + '>' +
        (interno ? '<i class="fas fa-download" aria-hidden="true"></i>&nbsp;Baixar' :
          '<i class="fas fa-external-link-alt" aria-hidden="true"></i>&nbsp;Acessar<span class="sr-only"> (abre em nova aba)</span>') + '</a>' : '') +
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
        (!ql || String((i.Titulo || '') + ' ' + (i.Descricao || '')).toLowerCase().indexOf(ql) >= 0);
    });
    var met = DADOS.params.Link_Metodologia, guia = DADOS.params.Link_Guia;
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>' + esc(par('Titulo_Repositorio', 'Repositório de materiais e ferramentas')) + '</h1><div class="linha" aria-hidden="true"></div></div>' +
      (met || guia ?
        '<div class="repo-oficial">' +
        (met ? '<a class="repo-oficial-card" href="' + esc(met) + '" target="_blank" rel="noopener"><i class="fas fa-scale-balanced" aria-hidden="true"></i><div><strong>Metodologia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicada na intranet/e-Codevasf</span></div><i class="fas fa-external-link-alt seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        (guia ? '<a class="repo-oficial-card" href="' + esc(guia) + '" target="_blank" rel="noopener"><i class="fas fa-book-open" aria-hidden="true"></i><div><strong>Guia de Gerenciamento de Processos</strong><span>RES 031/2025 · publicado na intranet/e-Codevasf</span></div><i class="fas fa-external-link-alt seta" aria-hidden="true"></i><span class="sr-only"> (abre em nova aba)</span></a>' : '') +
        '</div>' : '') +
      cardIniciativas() +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Jornada de mapeamento</h2><div class="linha" aria-hidden="true"></div></div>' +
      /* Componente Wizard: a jornada é um percurso linear e prescrito —
         painel de etapas (Step), área de conteúdo e barra de navegação com
         Cancelar terciário, Voltar secundário e Avançar primário (Concluir
         na última etapa). */
      '<div class="br-wizard" id="wizardJornada" step="1">' +
      '<div class="wizard-progress" role="tablist" aria-label="Etapas da jornada de mapeamento">' +
      DADOS.jornada.map(function (e, i) {
        return '<button class="wizard-progress-btn" type="button" role="tab" aria-controls="wj-' + i +
          '" aria-selected="false"><span class="info">' + esc(e.Nome) + '</span></button>';
      }).join('') + '</div>' +
      '<div class="wizard-form">' + DADOS.jornada.map(function (e, i) {
        return '<div class="wizard-panel" role="tabpanel" id="wj-' + i + '">' +
          '<div class="wizard-panel-content" tabindex="0">' +
          '<div class="jornada-etapa fase-' + slug(e.Fase || '') + '">' +
          '<div class="je-topo"><span class="je-num">' + esc(e.Ordem) + '</span><div><span class="je-fase">' + esc(e.Fase || '') + '</span><h4>' + esc(e.Nome) + '</h4></div><span class="je-dur"><i class="far fa-clock" aria-hidden="true"></i> ' + esc(e.Duracao || '') + '</span></div>' +
          '<p class="je-obj">' + esc(e.Objetivo || '') + '</p>' +
          '<div class="je-grid">' +
          '<div class="je-caixa"><b><i class="fas fa-list-check" aria-hidden="true"></i> Atividades-chave</b><ul>' + listar(e.Atividades_Chave).map(function (a) { return '<li>' + esc(a) + '</li>'; }).join('') + '</ul></div>' +
          '<div class="je-caixa"><b><i class="fas fa-people-group" aria-hidden="true"></i> Quem faz</b><p>' + esc(listar(e.Quem_Faz).join(' · ')) + '</p>' +
          '<b style="margin-top:8px"><i class="fas fa-box-open" aria-hidden="true"></i> Entregáveis</b><p>' + esc(listar(e.Entregaveis).join(' · ')) + '</p></div>' +
          '</div>' +
          (e.Sentimento_Usuario ? '<p class="je-sente"><i class="far fa-heart" aria-hidden="true"></i> ' + esc(e.Sentimento_Usuario) + '</p>' : '') +
          '</div></div>' +
          '<div class="wizard-panel-btn">' +
          '<button class="br-button wizard-btn-canc" type="button">Voltar ao início</button>' +
          '<button class="br-button secondary wizard-btn-prev" type="button">Voltar</button>' +
          '<button class="br-button primary wizard-btn-next" type="button">Avançar</button>' +
          '</div>' +
          '</div>';
      }).join('') + '</div></div></section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Instrumentos, modelos e ferramentas</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros de instrumentos, modelos e ferramentas">' +
      '<div class="filtros-campos">' +
      selectHtml({ chave: 'repoCat', id: 'repoCat', rotulo: 'Categoria',
        placeholder: 'Todas as categorias', selecionados: filtroRepo.cat ? [filtroRepo.cat] : [],
        opcoes: cats.map(function (c) { return { v: c, r: c }; }) }) +
      selectHtml({ chave: 'repoFase', id: 'repoFase', rotulo: 'Fase do ciclo',
        placeholder: 'Todas as fases', selecionados: filtroRepo.fase ? [filtroRepo.fase] : [],
        opcoes: fases.map(function (c) { return { v: c, r: c }; }) }) +
      buscaCampoHtml('repoQ', 'Pesquisar no repositório', 'Título, código ou descrição', filtroRepo.q) +
      '</div>' +
      rodapeFiltrosHtml('repo', [
        ['cat', filtroRepo.cat ? 'Categoria: ' + filtroRepo.cat : ''],
        ['fase', filtroRepo.fase ? 'Fase: ' + filtroRepo.fase : ''],
        ['q', filtroRepo.q ? 'Pesquisa: ' + filtroRepo.q : '']
      ]) +
      '</section>' +
      (lista.length ? '<div class="repo-grid">' + pagFatia('repo', lista, 6).map(cardRepo).join('') + '</div>' +
        paginacaoHtml('repo', lista.length, 'itens', [6, 12, 24, 48]) : vazio('Nenhum material com esses filtros',
          'Revise a categoria, a fase da jornada ou o termo buscado para ver os materiais publicados.',
          { img: 'empty-space/empty-space-04.png' })) + '</section>' +
      '<section class="pp-sec"><div class="pp-sec-h"><h2>Metodologia em resumo</h2><div class="linha" aria-hidden="true"></div></div>' +
      '<div class="pp-card"><h3><i class="fas fa-flag-checkered" aria-hidden="true"></i> Marcos do mapeamento (M1–M10)</h3>' +
      '<p style="font-size:var(--fs-sm);margin-bottom:var(--sp2)">Roteiro-padrão de cada projeto de mapeamento, do primeiro contato com a área até a publicação no repositório — abra o "i" de cada marco para ver o que ele significa:</p>' +
      '<ul class="marcos">' + MARCOS_ROTULOS.map(function (r, i) { return '<li class="feito"><span>' + esc(r) + '</span>' + dica(MARCOS_DESCRICOES[i]) + '<i class="fas fa-check-circle" aria-hidden="true"></i></li>'; }).join('') + '</ul>' +
      // Faixa contínua sob os marcos: monitoramento e avaliação não são um
      // marco, são atividades permanentes que atravessam todos eles.
      '<div class="marcos-continuo">' +
      '<span class="mc-rot"><i class="fas fa-arrows-rotate" aria-hidden="true"></i> Monitoramento e avaliação</span>' +
      '<span class="mc-trilho" aria-hidden="true"></span>' +
      '<span class="mc-nota">atividades contínuas — atravessam do M1 ao M10 e seguem depois da publicação</span>' +
      '</div></div>' +
      '<div class="br-message warning" role="status"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div><div class="content"><span class="message-title">Dados fictícios.</span> <span class="message-body">Todo o conteúdo exibido foi criado apenas para demonstrar o painel — substitua na planilha.</span></div><div class="close"><button class="br-button circle small" type="button" aria-label="Fechar a mensagem"><i class="fas fa-times" aria-hidden="true"></i></button></div></div></section>';
    ligarPaginacao(el, renderRepositorio);
    window.BRWizardInit(el);
    var wCanc = el.querySelector('#wizardJornada .wizard-btn-canc');
    if (wCanc) el.addEventListener('click', function (ev) {
      var b = ev.target.closest('#wizardJornada .wizard-btn-canc');
      if (!b) return;
      var primeira = el.querySelector('#wizardJornada .wizard-progress-btn');
      if (primeira) primeira.click();
    });
    ligarRodapeFiltros(el, 'repo', function (qual) {
      if (qual === 'cat' || qual === 'tudo') filtroRepo.cat = '';
      if (qual === 'fase' || qual === 'tudo') filtroRepo.fase = '';
      if (qual === 'q' || qual === 'tudo') filtroRepo.q = '';
      PAG.repo.pag = 1;
      renderRepositorio();
    });
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
      return String(p.Ponto_Focal_Nugep || '').toLowerCase().indexOf(alvo) === 0;
    });
  }
  /* Avatar (Componente Avatar, gov.br DS): a foto da planilha (coluna Foto)
     ocupa o lugar do avatar; as iniciais ficam por baixo e voltam a
     aparecer sozinhas se a URL da imagem falhar (onerror remove o <img>). */
  function avatarNugep(m, cls) {
    var foto = String(m.Foto || '').trim();
    return '<span class="br-avatar' + (cls ? ' ' + cls : '') + '" aria-hidden="true">' +
      '<span class="content">' + esc(iniciais(m.Nome)) +
      (foto ? '<img src="' + esc(foto) + '" alt="" loading="lazy" onerror="this.remove()">' : '') +
      '</span></span>';
  }
  function contatoNugep(m) {
    return '<div class="nugep-contato">' +
      (m.Email ? '<a href="mailto:' + esc(m.Email) + '"><i class="fas fa-envelope" aria-hidden="true"></i> ' + esc(m.Email) + '</a>' : '') +
      (m.Telefone ? '<a href="tel:+55' + esc(String(m.Telefone).replace(/\D/g, '')) + '"><i class="fas fa-phone" aria-hidden="true"></i> ' + esc(m.Telefone) + '</a>' : '') +
      '</div>';
  }
  // Perfil usado no bloco institucional: avatar + nome + papel + contatos.
  function perfilInst(m, cls) {
    return '<article class="perfil-inst' + (cls ? ' ' + cls : '') + '">' +
      avatarNugep(m) +
      '<div class="perfil-dados"><h3>' + esc(m.Nome) + '</h3>' +
      (m.Papel ? '<p class="perfil-papel">' + esc(m.Papel) + '</p>' : '') +
      contatoNugep(m) + '</div></article>';
  }
  function nivelNugep(m) { return +(m.Hierarquia || 0); }
  // Competências e atribuições (RES 031/2025, item 3) — instâncias da mais
  // alta à equipe do processo; accordion do gov.br DS, mesmo padrão do FAQ.
  function cardCompetencias() {
    var itens = DADOS.competencias || [];
    if (!itens.length) return '';
    return '<div class="pp-card" style="margin-top:var(--sp4)"><h3><i class="fas fa-scale-balanced" aria-hidden="true"></i> Competências e atribuições</h3>' +
      '<p class="pp-muted" style="margin-bottom:var(--sp2)">Atribuições de cada instância na gestão de processos, conforme a Resolução nº 031/2025 (item 3).</p>' +
      '<div class="br-accordion" id="comptAcc">' + itens.map(function (c, i) {
        var id = 'compt-' + i;
        return '<div class="item">' +
          '<button class="header" type="button" aria-controls="' + id + '" aria-expanded="false">' +
          '<span class="icon"><i class="fas fa-chevron-down" aria-hidden="true"></i></span>' +
          '<span class="title">' + esc(c.Instancia) + '</span></button></div>' +
          '<div class="content" id="' + id + '"><ul class="compt-lista">' +
          listar(c.Atribuicoes).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul></div>';
      }).join('') + '</div></div>';
  }
  function renderNugep() {
    var el = $('#viewNugep');
    var P = DADOS.params || {};
    /* Contato institucional na ordem hierárquica: Gerência-Executiva (1),
       Gerência (2) e, por último, a equipe da Unidade (3 — ou quem estiver
       lotado na AE/GPE/UNP, quando a planilha não traz a coluna). */
    var chefias = DADOS.nugep.filter(function (m) { return nivelNugep(m) === 1 || nivelNugep(m) === 2; })
      .sort(function (a, b) { return nivelNugep(a) - nivelNugep(b); });
    var equipeUnp = DADOS.nugep.filter(function (m) {
      return nivelNugep(m) === 3 || (!nivelNugep(m) && String(m.Unidade_Sigla || '') === 'AE/GPE/UNP');
    });
    var integrantes = DADOS.nugep.filter(function (m) { return nivelNugep(m) !== 1 && nivelNugep(m) !== 2; });
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>NUGEP — Núcleo de Gestão Normativa e de Processos</h1><div class="linha" aria-hidden="true"></div></div>' +
      (integrantes.length ? '<div class="nugep-grid">' + integrantes.map(function (m) {
        var meusProcs = processosDoNugep(m.Nome);
        return '<article class="nugep-card">' + avatarNugep(m) +
          '<h3>' + esc(m.Nome) + '</h3>' +
          '<p class="nugep-unid"><span class="nugep-sigla">' + siglaTag(m.Unidade_Sigla || '') + '</span></p>' +
          contatoNugep(m) +
          (meusProcs.length ? '<div class="nugep-procs"><b><i class="fas fa-diagram-project" aria-hidden="true"></i> Processos sob responsabilidade</b><ul>' +
            meusProcs.map(function (p) { return '<li><a href="#/p/' + encodeURIComponent(p.Trilha) + '">' + esc(p.Trilha) + ' — ' + esc(p.Nome) + '</a></li>'; }).join('') +
            '</ul></div>' : '') +
          '</article>';
      }).join('') + '</div>' : vazio('Nenhum integrante cadastrado',
        'A aba NUGEP da planilha ainda não tem integrantes. Assim que forem cadastrados, aparecem aqui.',
        { img: 'empty-space/empty-space-14.png' })) +
      '<div class="pp-card contato-inst" style="margin-top:var(--sp4)">' +
      '<h3><i class="fas fa-building" aria-hidden="true"></i> Contato institucional' +
      dica('Canal oficial da unidade e as pessoas responsáveis, na ordem hierárquica: gerência-executiva, gerência e a equipe da UNP.') +
      '</h3>' +
      '<p class="ci-unidade"><strong>' + esc(par('Contato_Unidade', 'Unidade de Gestão Normativa e de Processos (AE/GPE/UNP)')) + '</strong><br>' +
      (P.Contato_Email ? 'E-mail: <a href="mailto:' + esc(P.Contato_Email) + '">' + esc(P.Contato_Email) + '</a>' : '') +
      (P.Contato_Telefone ? ' · Telefone: ' + esc(P.Contato_Telefone) : '') + '</p>' +
      ((chefias.length || equipeUnp.length) ? '<div class="ci-hier">' +
        chefias.map(function (m) {
          return '<div class="ci-nivel"><span class="ci-rot">' + siglaTag(m.Unidade_Sigla || '') +
            (m.Unidade_Nome ? ' — ' + esc(m.Unidade_Nome) : '') + '</span>' +
            perfilInst(m, 'chefia') + '</div>';
        }).join('') +
        (equipeUnp.length ? '<div class="ci-nivel"><span class="ci-rot">AE/GPE/UNP — equipe da Unidade</span>' +
          '<div class="ci-equipe">' + equipeUnp.map(function (m) { return perfilInst(m); }).join('') + '</div></div>' : '') +
        '</div>' : '') +
      '</div>' +
      '<section class="pp-sec" id="sec-equipe-gestao"><div class="pp-sec-h"><h2>Equipe de Gerenciamento de Processos</h2><div class="linha" aria-hidden="true"></div></div>' +
      (DADOS.equipeGestao.length ? '<div class="nugep-grid">' + DADOS.equipeGestao.map(function (g) {
        return '<article class="nugep-card">' + avatarNugep(g) + '<h3>' + esc(g.Nome) + '</h3>' +
          (g.Unidade_Organica ? '<p class="nugep-unid"><span class="nugep-sigla">' + siglaTag(g.Unidade_Organica) + '</span></p>' : '') +
          contatoNugep(g) + '</article>';
      }).join('') + '</div>' : vazio('Nenhum gestor cadastrado',
        'A aba Equipe_Gerenciamento_Processos da planilha ainda não tem integrantes. Assim que forem cadastrados, aparecem aqui.',
        { img: 'empty-space/empty-space-14.png' })) +
      '</section>' +
      cardCompetencias();
    if (window.PPUI && PPUI.iniciarAccordions) PPUI.iniciarAccordions(el);
  }

  /* ── TELA: glossário ──────────────────────────────────────────────── */
  var filtroGloss = { q: '', letra: '', aba: 'termos' };
  var filtroSiglas = { q: '' };
  function renderGlossario() {
    var el = $('#viewGlossario');
    var ehSiglas = filtroGloss.aba === 'siglas';
    el.innerHTML =
      '<div class="pp-sec-h" style="margin-top:0"><h1>Glossário</h1><div class="linha" aria-hidden="true"></div></div>' +
      // Componente Tab oficial do DS (br-tab/tab-nav/tab-item — mesma anatomia
      // da faixa #navigation e das abas do dropdown de Alertas), não mais o
      // chip de filtro. Os dois botões apontam para o mesmo painel dinâmico
      // #glossCorpo (o conteúdo é regenerado ali, não existem dois tab-panels
      // separados) — simplificação intencional do padrão de abas, registrada
      // aqui por não haver conteúdo fixo por aba para dividir.
      '<div class="br-tab gloss-abas" id="glossTab">' +
      '<nav class="tab-nav" aria-label="Tipo de glossário"><ul role="tablist">' +
      '<li class="tab-item' + (!ehSiglas ? ' active' : '') + '" role="presentation"><button type="button" role="tab" id="glossTabTermos" aria-selected="' + (!ehSiglas) + '" aria-controls="glossCorpo" data-aba="termos"><i class="fas fa-spell-check" aria-hidden="true"></i> Termos de gestão de processos</button></li>' +
      '<li class="tab-item' + (ehSiglas ? ' active' : '') + '" role="presentation"><button type="button" role="tab" id="glossTabSiglas" aria-selected="' + ehSiglas + '" aria-controls="glossCorpo" data-aba="siglas"><i class="fas fa-building" aria-hidden="true"></i> Siglas das unidades</button></li>' +
      '</ul></nav></div>' +
      '<div id="glossCorpo" role="tabpanel" aria-labelledby="' + (ehSiglas ? 'glossTabSiglas' : 'glossTabTermos') + '"></div>';
    $all('.gloss-abas [data-aba]', el).forEach(function (b) {
      b.onclick = function () { filtroGloss.aba = b.getAttribute('data-aba'); renderGlossario(); };
    });
    if (ehSiglas) renderSiglasGlossario($('#glossCorpo')); else renderTermosGlossario($('#glossCorpo'));
  }
  function renderTermosGlossario(el) {
    var todos = DADOS.glossario;
    var ql = filtroGloss.q.toLowerCase();
    var lista = todos.filter(function (t) {
      var letra = String(t.Termo || '').charAt(0).toUpperCase();
      return (!filtroGloss.letra || letra === filtroGloss.letra) &&
        (!ql || String((t.Termo || '') + ' ' + (t.Definicao || '')).toLowerCase().indexOf(ql) >= 0);
    });
    var letrasDisp = {}; todos.forEach(function (t) { letrasDisp[String(t.Termo || '').charAt(0).toUpperCase()] = 1; });
    var abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var pagina = pagFatia('gloss', lista, 12);
    var porLetra = {};
    pagina.forEach(function (t) { var L = String(t.Termo || '').charAt(0).toUpperCase(); (porLetra[L] = porLetra[L] || []).push(t); });
    el.innerHTML =
      '<section class="pp-filtros-painel" role="search" aria-label="Filtros do glossário">' +
      '<div class="filtros-campos">' +
      buscaCampoHtml('glossQ', 'Pesquisar termo', 'Termo, sigla ou conceito (ex.: SIPOC, KPI, raia)', filtroGloss.q) +
      '</div>' +
      rodapeFiltrosHtml('gloss', [
        ['q', filtroGloss.q ? 'Pesquisa: ' + filtroGloss.q : ''],
        ['letra', filtroGloss.letra ? 'Letra: ' + filtroGloss.letra : '']
      ]) +
      '<span class="br-divider" role="presentation"></span>' +
      '<div class="gloss-abc" role="group" aria-label="Filtrar por letra"><button type="button" class="' + (filtroGloss.letra ? '' : 'ativo') + '" data-letra="">Todos</button>' +
      abc.map(function (L) { return '<button type="button" data-letra="' + L + '" class="' + (filtroGloss.letra === L ? 'ativo' : '') + '"' + (letrasDisp[L] ? '' : ' disabled') + '>' + L + '</button>'; }).join('') + '</div>' +
      '</section>' +
      (lista.length ? Object.keys(porLetra).sort().map(function (L) {
        return '<h3 class="gloss-letra">' + L + '</h3><div class="gloss-grid">' + porLetra[L].map(function (t) {
          return '<article class="gloss-card"><div class="gloss-topo"><h4>' + esc(t.Termo) + '</h4></div>' +
            '<p>' + esc(t.Definicao || '') + '</p>' +
            '<div class="gloss-rodape">' + (t.Fonte ? '<span class="repo-fonte">Fonte: ' + esc(t.Fonte) + '</span>' : '') +
            '</div></article>';
        }).join('') + '</div>';
      }).join('') : vazio('Nenhum termo encontrado',
        'Revise o termo buscado ou a letra selecionada para ver o glossário completo.')) +
      paginacaoHtml('gloss', lista.length, 'termos', [12, 24, 48, 96]);
    ligarPaginacao(el, function () { renderTermosGlossario(el); });
    var q = $('#glossQ');
    if (q) q.oninput = function () { filtroGloss.q = this.value; filtroGloss.letra = ''; PAG.gloss.pag = 1; renderTermosGlossario(el); var n = $('#glossQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
    $all('.gloss-abc button', el).forEach(function (b) { b.onclick = function () { filtroGloss.letra = b.getAttribute('data-letra'); renderTermosGlossario(el); }; });
    ligarRodapeFiltros(el, 'gloss', function (qual) {
      if (qual === 'q' || qual === 'tudo') filtroGloss.q = '';
      if (qual === 'letra' || qual === 'tudo') filtroGloss.letra = '';
      PAG.gloss.pag = 1;
      renderTermosGlossario(el);
    });
  }
  // Lista de Nomes e Siglas das Unidades Orgânicas (Decisão da Presidência
  // nº 601/2025, js/siglas.js) — aba irmã da de termos, mesmo padrão de busca.
  function renderSiglasGlossario(el) {
    var todas = DADOS.siglas || [];
    var ql = filtroSiglas.q.trim().toLowerCase();
    var lista = !ql ? todas : todas.filter(function (u) {
      return u.Sigla.toLowerCase().indexOf(ql) >= 0 || u.Nome.toLowerCase().indexOf(ql) >= 0;
    });
    var pagina = pagFatia('siglas', lista, 24);
    el.innerHTML =
      '<section class="pp-filtros-painel" role="search" aria-label="Filtro de siglas"><div class="filtros-campos">' +
      buscaCampoHtml('siglasQ', 'Pesquisar sigla ou unidade', 'Ex.: UNP, Gerência de Custos, AE/GFN', filtroSiglas.q) +
      '</div></section>' +
      (lista.length ? '<div class="siglas-grid">' + pagina.map(function (u) {
        return '<div class="siglas-item"><code class="siglas-cod">' + esc(u.Sigla) + '</code><span>' + esc(u.Nome) + '</span></div>';
      }).join('') + '</div>' : vazio('Nenhuma unidade encontrada', 'Revise o termo pesquisado.')) +
      paginacaoHtml('siglas', lista.length, 'unidades', [24, 48, 96]);
    ligarPaginacao(el, function () { renderSiglasGlossario(el); });
    var q = $('#siglasQ');
    if (q) q.oninput = function () { filtroSiglas.q = this.value; PAG.siglas.pag = 1; renderSiglasGlossario(el); var n = $('#siglasQ'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } };
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
      '<div class="pp-sec-h" style="margin-top:0"><h1>Perguntas frequentes</h1><div class="linha" aria-hidden="true"></div></div>' +
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
              '<span class="icon"><i class="fas fa-chevron-down" aria-hidden="true"></i></span>' +
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
      a: DADOS.ativs.filter(function (a) { return bate(a.Codigo) || bate(a.Nome); }),
      t: DADOS.tarefas.filter(function (t) { return bate(t.Codigo) || bate(t.Nome); }),
      doc: DADOS.docs.filter(function (x) { return bate(x.ID) || bate(x.Titulo); }),
      gl: DADOS.glossario.filter(function (t) { return bate(t.Termo) || bate(t.Definicao); }),
      rp: DADOS.repo.filter(function (i) { return bate(i.Titulo) || bate(i.Descricao); })
    };
    var total = r.mp.length + r.p.length + r.sp.length + r.a.length + r.doc.length + r.gl.length + r.rp.length + r.t.length;
    function linha(href, cod, nome, extra) {
      return '<div class="doc-item"><i class="fas fa-arrow-right fa-stack-ico" aria-hidden="true"></i><div>' +
        '<div class="tit"><a href="' + href + '"><span class="cod">' + esc(cod) + '</span> ' + esc(nome) + '</a></div>' +
        (extra ? '<div class="meta">' + extra + '</div>' : '') + '</div></div>';
    }
    el.innerHTML =
      breadcrumb([{ rotulo: 'Início', href: '#/' }, { rotulo: 'Busca' }]) +
      '<div class="pp-sec-h" style="margin-top:0"><h1>Resultados para “' + esc(q) + '”</h1><div class="linha" aria-hidden="true"></div></div>' +
      (total ? '' : vazio('Nada encontrado para “' + esc(q) + '”',
        'Tente outro termo, com menos palavras ou sem abreviações. Você também pode navegar pelo portfólio de processos.',
        { img: 'empty-space/empty-space-44.png',
          acoes: [{ rotulo: 'Ver o Portfólio', icone: 'fa-diagram-project', href: '#/catalogo' }] })) +
      grupo('Macroprocessos', r.mp, function (m) { return linha('#/mp/' + encodeURIComponent(m.Codigo), m._cod || m.Codigo, m.Nome, esc(m.Categoria)); }) +
      grupo('Processos', r.p, function (p) { return linha('#/p/' + encodeURIComponent(p.Trilha), codDisp(p.Codigo), p.Nome, esc(p.Status_Mapeamento) + ' · ' + p.Percentual + '%'); }) +
      grupo('Subprocessos', r.sp, function (s) { return linha('#/sp/' + encodeURIComponent(s.Trilha), codDisp(s.Codigo), s.Nome, ''); }) +
      grupo('Atividades', r.a, function (a) { return linha('#/a/' + encodeURIComponent(a.Trilha), codDisp(a.Codigo), a.Nome, ''); }) +
      grupo('Tarefas', r.t, function (t) { return linha('#/t/' + encodeURIComponent(t.Trilha), codDisp(t.Codigo), t.Nome, esc(t.Tipo_Tarefa || '')); }) +
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
    VISITAS_NAVEGADOR = contarVisitaRepositorio();
    var chip = $('#syncChip'); if (chip) chip.textContent = FONTE;
    var c;
    if ((c = $('#cntCatalogo'))) c.textContent = DADOS.procs.length;
    if ((c = $('#cntDocumentos'))) c.textContent = DADOS.docs.length;
    if ((c = $('#cntRiscos'))) c.textContent = DADOS.riscos.length;
    if ((c = $('#cntIndicadores'))) c.textContent = DADOS.metricas.length;
    if ((c = $('#cntRepositorio'))) c.textContent = DADOS.repo.length;
    if ((c = $('#cntNugep'))) c.textContent = DADOS.nugep.length;
    if ((c = $('#cntGlossario'))) c.textContent = DADOS.glossario.length;
    if ((c = $('#cntFaq'))) c.textContent = DADOS.faq.length;
    ligarAcoesCabecalho();
    montarNotificacoes();
    if (window.BRUploadInit) window.BRUploadInit();
    if (window.BRTooltipInit) window.BRTooltipInit();
    if (window.BRMessageInit) window.BRMessageInit();
    if (window.BRTableInit) window.BRTableInit();
    if (window.BRLoadingInit) window.BRLoadingInit();
    ajustarAbas();
    // Fontes e ícones mudam a largura das abas depois do primeiro layout:
    // recalcula quando a página termina de carregar.
    window.addEventListener('load', ajustarAbas);
    if (d.fonts && d.fonts.ready) d.fonts.ready.then(ajustarAbas);
    setTimeout(ajustarAbas, 400);
    var reAjuste;
    window.addEventListener('resize', function () {
      clearTimeout(reAjuste);
      reAjuste = setTimeout(ajustarAbas, 120);
    });
    if (window.PPUI) PPUI.setMenuSections([
      { rotulo: 'Início · Cadeia de Valor', href: '#/', icone: 'fa-house', meta: DADOS.macros.length + ' macro' },
      { rotulo: 'Catálogo de processos', href: '#/catalogo', icone: 'fa-layer-group', meta: DADOS.procs.length },
      { rotulo: 'Dashboard gerencial', href: '#/dashboard', icone: 'fa-chart-pie' },
      { rotulo: 'Repositório de materiais', href: '#/repositorio', icone: 'fa-toolbox' },
      { rotulo: 'Documentos', href: '#/documentos', icone: 'fa-folder-open', meta: DADOS.docs.length },
      { rotulo: 'Radar de riscos', href: '#/riscos', icone: 'fa-shield-halved', meta: DADOS.riscos.length },
      { rotulo: 'Indicadores', href: '#/indicadores', icone: 'fa-chart-line', meta: DADOS.metricas.length },
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
    }).join('') : '<div class="empty-state"><i class="fas fa-check-circle" aria-hidden="true"></i>Nenhum risco crítico aberto.</div>';
    painelP.innerHTML = prazos.length ? prazos.map(function (p) {
      return item('warning', p.Nome, 'Prazo em ' + fmtData(p.Prazo_Previsto),
        codDisp(p.Codigo) + ' · ' + (p.Status_Mapeamento || '') + ' · ' + pctNorm(p.Percentual) + '%',
        '#/p/' + encodeURIComponent(p.Trilha));
    }).join('') : '<div class="empty-state"><i class="fas fa-check-circle" aria-hidden="true"></i>Nenhum mapeamento com prazo vencido.</div>';
    var total = riscos.length + prazos.length;
    var cR = $('#notifCountRiscos'), cP = $('#notifCountPrazos'), badge = $('#notifBadge');
    if (cR) cR.textContent = riscos.length;
    if (cP) cP.textContent = prazos.length;
    if (badge) { badge.textContent = total; badge.hidden = total === 0; }
  }

  function iniciar() {
    var v = $('#viewInicio');
    /* Padrão Skeleton Screen: em vez do giro sem contexto, o painel
       desenha a forma da tela inicial (título, faixa de indicadores e
       cartões) enquanto a planilha carrega. O status em aria-live continua
       a cargo do leitor de tela — o esqueleto é decorativo. */
    if (v) v.innerHTML = '<div class="pp-loading br-skeleton" aria-hidden="true">' +
      '<div class="skeleton-group" style="max-width:520px;margin-bottom:var(--sp5)">' +
      '<div class="skeleton-title lg" style="width:70%"></div>' +
      '<div class="skeleton-line w-90"></div>' +
      '<div class="skeleton-line w-50"></div></div>' +
      '<div style="display:grid;gap:var(--sp2);grid-template-columns:repeat(auto-fit,minmax(160px,1fr));margin-bottom:var(--sp5)">' +
      '<div class="skeleton-block" style="height:104px"></div>'.repeat(5) + '</div>' +
      '<div style="display:grid;gap:var(--sp3);grid-template-columns:repeat(auto-fit,minmax(260px,1fr))">' +
      '<div class="skeleton-block" style="height:180px"></div>'.repeat(3) + '</div></div>' +
      '<p class="loading-label" role="status" aria-live="polite">Carregando dados do painel…</p>';
    carregarDados().then(posCarga).catch(function (e) {
      console.error(e);
      /* Tela de erro ilustrada (Fundamento Ilustração > Cenários > Erro):
         a ilustração suaviza a falha e o texto diz o que fazer. */
      if (v) v.innerHTML = '<div class="pp-erro-ilus"><img src="img/ilustracoes/erro/error10.png" alt="" aria-hidden="true"></div>' +
        '<div class="br-message warning" role="alert"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div><div class="content"><span class="message-title">Não foi possível carregar os dados.</span> <span class="message-body">Verifique se data/painel-processos-dados.xlsx está publicado (ou gere js/dados.js com scripts/planilha_para_js.py). Detalhe: ' + esc(e.message) + '</span></div></div>';
    });
  }
  window.addEventListener('hashchange', rota);
  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
