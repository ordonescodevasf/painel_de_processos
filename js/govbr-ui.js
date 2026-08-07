/* ════════════════════════════════════════════════════════════════════
   GOVBR-UI — comportamentos genéricos do cabeçalho, menu, busca,
   contraste, rodapé e modais (gov.br DS v4), reimplementados a partir
   do Painel de Transformação Digital para uso neste painel.
   Sem dependências externas.
   ════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';
  var d = document;

  function $(sel, ctx) { return (ctx || d).querySelector(sel); }
  function $all(sel, ctx) { return Array.prototype.slice.call((ctx || d).querySelectorAll(sel)); }

  /* ── DROPDOWNS (data-toggle="dropdown" data-target="id") ─────────── */
  function closeAllDropdowns(except) {
    $all('.dd-target').forEach(function (p) {
      if (p !== except && !p.hidden) {
        p.hidden = true;
        var b = $('[data-target="' + p.id + '"]');
        if (b) b.setAttribute('aria-expanded', 'false');
      }
    });
  }
  /* Dropdowns marcados "dd-fixed" (ex.: o "Mais" da barra de abas, que vive
     dentro de #navigation, position:sticky) são posicionados com
     position:fixed calculado em JS a partir do botão-gatilho, em vez de
     absolute/relative — assim ficam imunes a qualquer contexto de
     empilhamento ou corte de overflow criado por ancestrais (barra de
     abas fixa, cartões com z-index local etc.), sempre por cima. */
  function posicionarDropdownFixo(panel, btn) {
    var r = btn.getBoundingClientRect();
    var margem = 8;
    panel.style.position = 'fixed';
    panel.style.top = Math.round(r.bottom + 6) + 'px';
    panel.style.left = 'auto';
    var direita = window.innerWidth - r.right;
    if (direita < margem) direita = margem;
    panel.style.right = Math.round(direita) + 'px';
    panel.style.maxWidth = 'calc(100vw - ' + (margem * 2) + 'px)';
    // se a largura mínima do painel não couber à esquerda do gatilho, gruda na margem esquerda
    var larguraMin = parseFloat(getComputedStyle(panel).minWidth) || 0;
    if (r.right - larguraMin < margem) { panel.style.right = margem + 'px'; }
  }
  function fecharDropdownsFixosNoScroll() {
    $all('.dd-target.dd-fixed').forEach(function (p) {
      if (!p.hidden) {
        p.hidden = true;
        var b = $('[data-target="' + p.id + '"]');
        if (b) b.setAttribute('aria-expanded', 'false');
      }
    });
  }
  window.addEventListener('scroll', fecharDropdownsFixosNoScroll, { passive: true, capture: true });
  window.addEventListener('resize', fecharDropdownsFixosNoScroll);

  d.addEventListener('click', function (ev) {
    var btn = ev.target.closest('[data-toggle="dropdown"]');
    if (btn) {
      var panel = d.getElementById(btn.getAttribute('data-target'));
      if (!panel) return;
      var abrir = panel.hidden;
      closeAllDropdowns(abrir ? panel : null);
      panel.hidden = !abrir;
      btn.setAttribute('aria-expanded', abrir ? 'true' : 'false');
      if (abrir && panel.classList.contains('dd-fixed')) posicionarDropdownFixo(panel, btn);
      ev.stopPropagation();
      return;
    }
    if (!ev.target.closest('.dd-target')) closeAllDropdowns(null);
  });

  /* ── ACCORDION (br-accordion) — Componentes > Accordion, gov.br DS ──
     Porte da classe BRAccordion publicada: alterna o atributo `active` do
     item clicado e troca o ícone angle-down/angle-up. Com o atributo
     `single` na raiz, abrir um item fecha os demais; sem ele, vários podem
     ficar abertos ao mesmo tempo (comportamento padrão). aria-expanded
     acompanha o estado, para o leitor de tela. ── */
  function BRAccordion(nome, componente) {
    this.name = nome;
    this.component = componente;
    this._setBehavior();
  }
  BRAccordion.prototype._setBehavior = function () {
    var self = this;
    Array.prototype.forEach.call(this.component.querySelectorAll('button.header'), function (botao) {
      botao.addEventListener('click', function (ev) {
        self._collapse(ev);
        self._changeIcon();
      });
    });
  };
  BRAccordion.prototype._collapse = function (ev) {
    var unico = this.component.hasAttribute('single');
    var alvo = ev.currentTarget.parentNode;
    Array.prototype.forEach.call(this.component.querySelectorAll('.item'), function (item) {
      if (item === alvo) {
        if (item.hasAttribute('active')) item.removeAttribute('active');
        else item.setAttribute('active', '');
      } else if (unico && item.hasAttribute('active')) {
        item.removeAttribute('active');
      }
    });
  };
  BRAccordion.prototype._changeIcon = function () {
    Array.prototype.forEach.call(this.component.querySelectorAll('.item'), function (item) {
      var aberto = item.hasAttribute('active');
      Array.prototype.forEach.call(item.querySelectorAll('.icon'), function (ic) {
        var i = ic.children[0];
        if (!i) return;
        i.classList.toggle('fa-angle-up', aberto);
        i.classList.toggle('fa-angle-down', !aberto);
      });
      var h = item.querySelector('button.header');
      if (h) h.setAttribute('aria-expanded', String(aberto));
    });
  };
  // Instancia todos os accordions ainda não ligados dentro de `escopo`.
  function iniciarAccordions(escopo) {
    var raiz = escopo || d;
    Array.prototype.forEach.call(raiz.querySelectorAll('.br-accordion'), function (el) {
      if (el.__brAccordion) return;
      el.__brAccordion = new BRAccordion('br-accordion', el);
    });
  }

  /* ── COLLAPSE (data-toggle="collapse" data-target="id") ──────────── */
  d.addEventListener('click', function (ev) {
    var btn = ev.target.closest('[data-toggle="collapse"]');
    if (!btn) return;
    var alvo = d.getElementById(btn.getAttribute('data-target'));
    if (!alvo) return;
    var abrir = alvo.hidden;
    alvo.hidden = !abrir;
    alvo.setAttribute('aria-hidden', abrir ? 'false' : 'true');
    btn.setAttribute('aria-expanded', abrir ? 'true' : 'false');
    var ic = btn.querySelector('.fa-chevron-up,.fa-chevron-down');
    if (ic) { ic.classList.toggle('fa-chevron-up', abrir); ic.classList.toggle('fa-chevron-down', !abrir); }
  });

  /* ── MENU LATERAL (data-toggle="menu" / data-dismiss="menu") ─────── */
  var menuTriggerAtual = null;
  d.addEventListener('click', function (ev) {
    var abre = ev.target.closest('[data-toggle="menu"]');
    var fecha = ev.target.closest('[data-dismiss="menu"]');
    var menu = $('#sectionMenu');
    if (!menu) return;
    if (abre) {
      menu.classList.add('active');
      abre.setAttribute('aria-expanded', 'true');
      menuTriggerAtual = abre;
      var f = menu.querySelector('a,button'); if (f) f.focus();
    } else if (fecha || (menu.classList.contains('active') && !ev.target.closest('.menu-panel') && ev.target.closest('#sectionMenu'))) {
      menu.classList.remove('active');
      if (menuTriggerAtual) { menuTriggerAtual.setAttribute('aria-expanded', 'false'); menuTriggerAtual.focus(); }
    }
  });
  d.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape') return;
    var menu = $('#sectionMenu.active');
    if (menu) { menu.classList.remove('active'); if (menuTriggerAtual) menuTriggerAtual.focus(); }
    closeAllDropdowns(null);
    $all('.br-scrim.foco:not([hidden])').forEach(function (s) { s.hidden = true; });
    var hs = $('#headerSearch'); if (hs) hs.classList.remove('active');
  });

  /* ── Submenu gov.br dentro do menu lateral (setas / voltar) ──────── */
  d.addEventListener('click', function (ev) {
    var seta = ev.target.closest('#sectionMenu .br-button.arrow');
    if (seta) {
      var off = seta.parentElement.querySelector(':scope > div.off');
      if (off) { off.hidden = false; off.setAttribute('aria-hidden', 'false'); seta.setAttribute('aria-expanded', 'true'); }
      ev.preventDefault(); return;
    }
    var voltar = ev.target.closest('#sectionMenu .backButton');
    if (voltar) {
      var painel = voltar.closest('div.off');
      if (painel) { painel.hidden = true; painel.setAttribute('aria-hidden', 'true'); }
      ev.preventDefault();
    }
  });

  /* ── BUSCA do cabeçalho (data-toggle="search" / data-dismiss) ────── */
  d.addEventListener('click', function (ev) {
    if (ev.target.closest('[data-toggle="search"]')) {
      var box = $('#headerSearch');
      if (box) { box.classList.add('active'); var i = $('#headerSearchInput'); if (i) i.focus(); }
    }
    if (ev.target.closest('[data-dismiss="search"]')) {
      var b2 = $('#headerSearch'); if (b2) b2.classList.remove('active');
    }
  });

  /* ── CONTRASTE (persistente) ─────────────────────────────────────── */
  function aplicarContraste(on) {
    d.body.classList.toggle('contraste-alto', on);
    var btn = $('#btnContrastIcon');
    if (btn) {
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
      btn.setAttribute('aria-label', on ? 'Desativar contraste' : 'Ativar contraste');
    }
    try { localStorage.setItem('pp-contraste', on ? '1' : '0'); } catch (e) {}
  }
  d.addEventListener('click', function (ev) {
    if (ev.target.closest('#btnContrastIcon')) {
      aplicarContraste(!d.body.classList.contains('contraste-alto'));
    }
  });
  d.addEventListener('DOMContentLoaded', function () {
    var pref = null;
    try { pref = localStorage.getItem('pp-contraste'); } catch (e) {}
    if (pref === '1') aplicarContraste(true);
  });

  /* ── MODAIS (.br-scrim.foco) — [data-modal-open="id"] abre;
        [data-dismiss="true"] ou clique no scrim fecha ─────────────── */
  d.addEventListener('click', function (ev) {
    var abre = ev.target.closest('[data-modal-open]');
    if (abre) {
      ev.preventDefault();
      var m = d.getElementById(abre.getAttribute('data-modal-open'));
      if (m) { m.hidden = false; var f = m.querySelector('.br-modal'); if (f) f.focus && f.focus(); }
      return;
    }
    var fecha = ev.target.closest('[data-dismiss="true"]');
    if (fecha) { var s = fecha.closest('.br-scrim'); if (s) s.hidden = true; return; }
    if (ev.target.classList && ev.target.classList.contains('br-scrim')) ev.target.hidden = true;
  });

  /* ── Cookies: o botão da barra superior reabre o cookiebar na etapa
        "aberto", onde as classes de cookies podem ser reconfiguradas ── */

  /* ── RODAPÉ — acordeão do mapa do site (data-fs-toggle) ──────────── */
  d.addEventListener('click', function (ev) {
    var btn = ev.target.closest('[data-fs-toggle]');
    if (!btn) return;
    var col = btn.closest('.fs-col');
    var itens = col && col.querySelector('.fs-items');
    if (!itens) return;
    var abrir = itens.hidden;
    // comportamento oficial: apenas uma categoria aberta por vez (mobile)
    $all('#footerSitemap .fs-col').forEach(function (c) {
      var it = c.querySelector('.fs-items'); var b = c.querySelector('[data-fs-toggle]');
      var ic = b && b.querySelector('i');
      var abre = (c === col) && abrir;
      if (it) it.hidden = !abre;
      if (b) b.setAttribute('aria-expanded', abre ? 'true' : 'false');
      if (ic) { ic.classList.toggle('fa-angle-up', abre); ic.classList.toggle('fa-angle-down', !abre); }
    });
  });

  /* ── VOLTAR AO TOPO — some/aparece conforme a rolagem ─────────────── */
  (function () {
    var btn = d.getElementById('btnBackToTop');
    if (!btn) return;
    function syncVisibility() {
      if (window.scrollY > 400) btn.removeAttribute('hidden');
      else btn.setAttribute('hidden', '');
    }
    window.addEventListener('scroll', syncVisibility, { passive: true });
    syncVisibility();
    btn.addEventListener('click', function () {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  })();

  /* ── REPORTAR ERRO — tenta capturar print da tela (Screen Capture API,
     exige permissão do usuário — não existe captura silenciosa) e abre
     um e-mail com a descrição; como mailto: não aceita anexos, o print
     é oferecido para download com instrução de anexar manualmente. ── */
  (function () {
    var openBtn = d.getElementById('btnReportError');
    var scrimEl = d.getElementById('reportErrorScrim');
    if (!openBtn || !scrimEl) return;
    var retryBtn = d.getElementById('btnRetryShot');
    var sendBtn = d.getElementById('btnSendReport');
    var statusEl = d.getElementById('reportErrorShotStatus');
    var previewEl = d.getElementById('reportErrorShotPreview');
    // Figura com Scrim Legibilidade: a legenda fica sobre o print.
    var figureEl = d.getElementById('reportErrorShotFigure');
    var downloadEl = d.getElementById('reportErrorShotDownload');
    var textEl = d.getElementById('reportErrorText');
    var shotDataUrl = null;

    function setStatus(msg, carregando) {
      statusEl.textContent = msg;
      // Modal carregando (dependência Loading do Componente Modal): o
      // indicador só aparece enquanto a captura está em andamento.
      var load = d.getElementById('reportErrorShotLoading');
      if (load) load.hidden = !carregando;
    }

    function captureScreenshot() {
      shotDataUrl = null;
      figureEl.setAttribute('hidden', '');
      downloadEl.setAttribute('hidden', '');
      if (!navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) {
        setStatus('Este navegador não permite captura automática de tela. Tire um print manualmente (Windows: tecla Print Screen ou Win+Shift+S; Mac: Cmd+Shift+4) e anexe ao e-mail.');
        return;
      }
      setStatus('Capturando print da tela… escolha a tela/aba na janela que o navegador abrir.', true);
      navigator.mediaDevices.getDisplayMedia({ video: true }).then(function (stream) {
        var track = stream.getVideoTracks()[0];
        var video = d.createElement('video');
        video.srcObject = stream;
        video.muted = true;
        return video.play().then(function () {
          return new Promise(function (resolve) { setTimeout(resolve, 250); });
        }).then(function () {
          var canvas = d.createElement('canvas');
          canvas.width = video.videoWidth || 1280;
          canvas.height = video.videoHeight || 720;
          canvas.getContext('2d').drawImage(video, 0, 0, canvas.width, canvas.height);
          track.stop();
          shotDataUrl = canvas.toDataURL('image/png');
          previewEl.src = shotDataUrl;
          figureEl.removeAttribute('hidden');
          downloadEl.href = shotDataUrl;
          downloadEl.removeAttribute('hidden');
          setStatus('Print capturado. Baixe-o e anexe ao e-mail antes de enviar.');
        });
      }).catch(function () {
        setStatus('Captura cancelada ou não permitida. Tire um print manualmente (Windows: tecla Print Screen ou Win+Shift+S; Mac: Cmd+Shift+4) e anexe ao e-mail, ou tente novamente.');
      });
    }

    openBtn.addEventListener('click', function () {
      textEl.value = '';
      scrimEl.hidden = false;
      captureScreenshot();
    });
    if (retryBtn) retryBtn.addEventListener('click', captureScreenshot);

    if (sendBtn) sendBtn.addEventListener('click', function (ev) {
      ev.preventDefault();
      var desc = textEl.value.trim() || '(não informada)';
      var secao = (d.querySelector('#mainTabContent .tab-panel.active h1, #mainTabContent .tab-panel.active h2') || {}).textContent || d.title;
      var lines = [
        'Descrição do problema:', desc, '',
        'Seção ativa: ' + secao.trim() + ' (' + (location.hash || '#/') + ')',
        'Navegador: ' + navigator.userAgent,
        'Data/hora: ' + new Date().toLocaleString('pt-BR')
      ];
      if (shotDataUrl) lines.push('', 'IMPORTANTE: anexe a este e-mail o print que foi baixado (botão "Baixar print" no painel) antes de enviar.');
      var subject = encodeURIComponent('Erro no Painel de Processos');
      var body = encodeURIComponent(lines.join('\n'));
      window.location.href = 'mailto:ae.gpe.unp@codevasf.gov.br?subject=' + subject + '&body=' + body;
    });
  })();

  /* ── COOKIEBAR (br-cookiebar) — Componente gov.br DS 3.7.0 completo:
     etapa default (texto informativo + área de ação) e etapa aberta
     (conteúdo introdutório, principal e complementar), construído a
     partir de um JSON de entrada e devolvendo um JSON de saída no
     aceite, como no core oficial. Padrão opt out: as classes opcionais
     podem ser desligadas por classe (checkbox) ou por recurso (switch),
     e a escolha é de fato respeitada pelo painel. ── */
  (function () {
    var KEY = 'painel_processos_cookie_consent_v2';
    var comp = d.getElementById('cookieNotice');
    if (!comp) return;
    var scrim = d.getElementById('cookieScrim');

    var DADOS = {
      lang: 'pt-br',
      allOptOut: true,
      acceptButton: 'Aceitar',
      optOutButton: 'Definir cookies',
      closeLabel: 'Fechar as definições de cookies',
      infoText: 'Este painel carrega recursos de terceiros para funcionar (fontes, ícones, leitura da planilha e tradução para Libras) e guarda no seu navegador apenas as suas preferências de uso. Não coletamos dados pessoais. Você pode aceitar a configuração padrão ou definir quais recursos opcionais deseja usar.',
      mainTitle: 'Recursos e dados usados por este painel',
      lastUpdate: '06/08/2026',
      entryText: 'Abaixo estão todos os recursos externos que o painel carrega e tudo o que ele guarda no seu navegador, com finalidade, prazo e responsável. Os recursos essenciais não podem ser desligados porque sem eles o painel não carrega. Os opcionais ficam sob seu controle e a escolha vale a partir do próximo carregamento.',
      selectAll: true,
      allAlertMessage: 'Com recursos opcionais desligados, algumas funcionalidades do painel deixam de funcionar.',
      cookieGroups: [
        {
          groupId: 'essenciais',
          groupName: 'Estritamente necessários',
          groupOptOut: false,
          groupSelected: true,
          groupText: 'Recursos sem os quais o painel não carrega ou perde funções básicas de leitura e acessibilidade.',
          cookieList: [
            {
              cookieId: 'google-fonts', cookieOptOut: false, cookieSelected: true,
              cookieName: 'fonts.googleapis.com / fonts.gstatic.com', expires: 'Cache do navegador',
              domain: 'gstatic.com', entreprise: 'Google LLC',
              purpose: 'Carregar a tipografia oficial do padrão gov.br (Noto Sans e Noto Sans Mono).',
              description: 'Requisição de arquivos de fonte. O serviço pode registrar dados técnicos de acesso, como endereço IP, segundo a política do próprio provedor.'
            },
            {
              cookieId: 'cdn-cloudflare', cookieOptOut: false, cookieSelected: true,
              cookieName: 'cdnjs.cloudflare.com', expires: 'Cache do navegador',
              domain: 'cloudflare.com', entreprise: 'Cloudflare, Inc.',
              purpose: 'Ícones (Font Awesome) e biblioteca de leitura da planilha .xlsx (SheetJS).',
              description: 'Sem esses arquivos o painel não exibe ícones nem consegue ler a base de dados local.'
            },
            {
              cookieId: 'preferencias-locais', cookieOptOut: false, cookieSelected: true,
              cookieName: 'Preferências locais (localStorage)', expires: 'Até limpar os dados do navegador',
              domain: 'Este navegador', entreprise: 'Codevasf',
              purpose: 'Guardar alto contraste, avisos já lidos e a dispensa do tutorial.',
              description: 'Ficam só no seu equipamento e nunca são enviadas a nenhum servidor.'
            }
          ]
        },
        {
          groupId: 'opcionais',
          groupName: 'Recursos opcionais',
          groupOptOut: true,
          groupSelected: true,
          groupAlertMessage: 'Esta classe está parcial ou totalmente desligada — os recursos correspondentes não serão carregados.',
          groupText: 'Recursos que ampliam o painel, mas que você pode dispensar sem perder o acesso ao conteúdo.',
          cookieList: [
            {
              cookieId: 'vlibras', cookieOptOut: true, cookieSelected: true,
              alertMessage: 'Sem o VLibras o painel deixa de oferecer tradução automática para Libras.',
              cookieName: 'vlibras.gov.br', expires: 'Sessão',
              domain: 'vlibras.gov.br', entreprise: 'Governo Federal',
              purpose: 'Widget oficial de tradução do conteúdo para Libras.',
              description: 'Carrega o plugin do VLibras e o exibe no canto da tela.'
            },
            {
              cookieId: 'google-sheets', cookieOptOut: true, cookieSelected: true,
              alertMessage: 'Sem a leitura remota o painel usa apenas a planilha local e pode ficar desatualizado.',
              cookieName: 'docs.google.com / sheets.googleapis.com', expires: 'Sessão',
              domain: 'google.com', entreprise: 'Google LLC',
              purpose: 'Ler a base de dados publicada em planilha, quando o painel estiver configurado assim.',
              description: 'Somente leitura de dados públicos de processos. Nenhum dado seu é enviado.'
            }
          ]
        }
      ],
      noteTitle: 'Aviso sobre cookies',
      noteList: [
        { question: 'Este painel usa cookies de rastreio?', answer: 'Não. O painel não define cookies próprios nem usa ferramentas de análise de audiência. O que existe são requisições a serviços externos e preferências guardadas no seu navegador.' },
        { question: 'O que acontece se eu recusar os recursos opcionais?', answer: 'O painel continua funcionando por completo na parte de conteúdo. Você perde apenas a tradução para Libras e a atualização remota da base, que passa a vir da planilha local.' },
        { question: 'Como mudo a escolha depois?', answer: 'Pelo ícone de cookies na barra superior do painel, a qualquer momento.' }
      ],
      links: [
        { name: 'Lei nº 13.709/2018 (LGPD)', url: 'https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/l13709.htm', target: '_blank' },
        { name: 'Política de privacidade do gov.br', url: 'https://www.gov.br/pt-br/termo-de-uso-e-aviso-de-privacidade', target: '_blank' },
        { name: 'Falar com a UNP', url: 'mailto:ae.gpe.unp@codevasf.gov.br', target: '_self' }
      ]
    };

    function esc(t) {
      return String(t == null ? '' : t).replace(/[&<>"]/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
      });
    }
    function todosCookies() {
      var t = [];
      DADOS.cookieGroups.forEach(function (g) { g.cookieList.forEach(function (c) { t.push(c); }); });
      return t;
    }
    function selecionados() {
      return todosCookies().filter(function (c) { return c.cookieSelected; }).length;
    }

    function grupoHtml(g, gi) {
      var optOut = g.groupOptOut;
      var controle = optOut
        ? '<div class="br-checkbox"><input id="ck-group-' + gi + '" name="ck-group-' + gi + '" type="checkbox" data-group="' + gi + '"' + (g.groupSelected ? ' checked' : '') + '><label for="ck-group-' + gi + '">Selecionar toda a classe</label></div>'
        : '<span class="always-active">Sempre ativo</span>';
      return '<div class="group-block" data-group-block="' + gi + '">' +
        '<div class="group-info">' +
        '<button class="group-name" type="button" data-toggle-group="' + gi + '" aria-expanded="false" aria-controls="cookie-list-' + gi + '">' +
        esc(g.groupName) + ' <span class="group-size">(' + g.cookieList.length + ')</span></button>' +
        '<div class="group-actions">' + controle +
        '<button class="br-button circle small terciary" type="button" data-toggle-group="' + gi + '" aria-label="Expandir a classe ' + esc(g.groupName) + '" aria-expanded="false" aria-controls="cookie-list-' + gi + '"><i class="fas fa-angle-down" aria-hidden="true"></i></button>' +
        '</div>' +
        '<p class="group-description">' + esc(g.groupText) + '</p>' +
        (g.groupAlertMessage ? '<div class="br-message warning d-none" role="alert" data-group-alert="' + gi + '"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div><div class="content"><span>' + esc(g.groupAlertMessage) + '</span></div></div>' : '') +
        '</div>' +
        '<div class="cookie-list' + (optOut ? '' : ' br-scrim inibicao') + '" id="cookie-list-' + gi + '">' +
        g.cookieList.map(function (c, ci) { return cookieHtml(g, gi, c, ci); }).join('') +
        '</div></div>';
    }
    function cookieHtml(g, gi, c, ci) {
      var sw = c.cookieOptOut
        ? '<div class="br-switch small right"><input id="sw-cookie-' + gi + '-' + ci + '" type="checkbox" role="switch" data-cookie="' + gi + '-' + ci + '"' + (c.cookieSelected ? ' checked' : '') + '><label for="sw-cookie-' + gi + '-' + ci + '">' + (c.cookieSelected ? 'Ligado' : 'Desligado') + '</label></div>'
        : '<span class="always-active">Sempre ativo</span>';
      return '<div class="cookie-info"><div class="br-card"><div class="card-content">' +
        '<div class="cookie-head"><span class="cookie-term">' + esc(c.cookieName) + '</span>' + sw + '</div>' +
        '<dl class="cookie-props">' +
        '<dt>Vencimento</dt><dd>' + esc(c.expires) + '</dd>' +
        '<dt>Domínio</dt><dd>' + esc(c.domain) + '</dd>' +
        '<dt>Empresa</dt><dd>' + esc(c.entreprise) + '</dd>' +
        '<dt>Finalidade</dt><dd>' + esc(c.purpose) + '</dd>' +
        '<dt>Descrição</dt><dd>' + esc(c.description) + '</dd>' +
        '</dl>' +
        (c.alertMessage ? '<div class="br-message warning d-none" role="alert" data-cookie-alert="' + gi + '-' + ci + '"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div><div class="content"><span>' + esc(c.alertMessage) + '</span></div></div>' : '') +
        '</div></div></div>';
    }

    function construir() {
      comp.setAttribute('role', 'dialog');
      comp.setAttribute('aria-modal', 'true');
      comp.setAttribute('aria-describedby', 'cookiebar-info');
      comp.setAttribute('aria-label', 'Componente para definição de cookies');
      comp.innerHTML =
        '<div class="br-modal"><div class="wrapper">' +
        '<div class="br-modal-header entry-content">' +
        '<h2 class="br-modal-title">' + esc(DADOS.mainTitle) + '</h2>' +
        '<button class="br-button circle terciary close" type="button" id="cookiebarClose" aria-label="' + esc(DADOS.closeLabel) + '"><i class="fas fa-times" aria-hidden="true"></i></button>' +
        '<p class="last-update">Última atualização: <span>' + esc(DADOS.lastUpdate) + '</span></p>' +
        '<p class="entry-text">' + esc(DADOS.entryText) + '</p>' +
        '</div>' +
        '<p class="info-text" id="cookiebar-info">' + esc(DADOS.infoText) + '</p>' +
        '<div class="br-modal-body main-content">' +
        '<h3 class="cookie-groups-title">Classes de cookies</h3>' +
        '<div class="select-all-bar">' +
        '<div class="br-checkbox"><input id="ck-all" name="ck-all" type="checkbox" checked><label for="ck-all">Selecionar tudo</label></div>' +
        '<span class="cookies-checked" id="cookiesChecked"></span>' +
        '</div>' +
        '<div class="br-message warning d-none" role="alert" id="cookiebarAllAlert"><div class="icon"><i class="fas fa-exclamation-triangle" aria-hidden="true"></i></div><div class="content"><span>' + esc(DADOS.allAlertMessage) + '</span></div></div>' +
        '<div class="br-list">' + DADOS.cookieGroups.map(grupoHtml).join('') + '</div>' +
        '<div class="notes"><h3>' + esc(DADOS.noteTitle) + '</h3><dl>' +
        DADOS.noteList.map(function (n) { return '<dt>' + esc(n.question) + '</dt><dd>' + esc(n.answer) + '</dd>'; }).join('') +
        '</dl></div>' +
        '</div>' +
        '<div class="complementary-content">' +
        DADOS.links.map(function (l) {
          return '<a href="' + esc(l.url) + '" target="' + esc(l.target || '_self') + '"' + (l.target === '_blank' ? ' rel="noopener"' : '') + '>' + esc(l.name) + '</a>';
        }).join('') +
        '</div>' +
        '</div>' +
        '<div class="br-modal-footer actions">' +
        '<button class="br-button secondary" type="button" id="cookiebarPolitics">' + esc(DADOS.optOutButton) + '</button>' +
        '<button class="br-button primary" type="button" id="cookiebarAccept">' + esc(DADOS.acceptButton) + '</button>' +
        '</div></div>';
    }

    function sincronizar() {
      var total = todosCookies().length, sel = selecionados();
      var todos = comp.querySelector('#ck-all');
      todos.checked = sel > 0;
      todos.indeterminate = sel > 0 && sel < total;
      comp.querySelector('#cookiesChecked').textContent = sel + ' de ' + total + ' recursos selecionados';
      var alerta = comp.querySelector('#cookiebarAllAlert');
      alerta.classList.toggle('d-none', sel === total);
      DADOS.cookieGroups.forEach(function (g, gi) {
        var opcionais = g.cookieList.filter(function (c) { return c.cookieOptOut; });
        var marcados = opcionais.filter(function (c) { return c.cookieSelected; }).length;
        var ck = comp.querySelector('#ck-group-' + gi);
        if (ck) {
          ck.checked = marcados > 0;
          ck.indeterminate = marcados > 0 && marcados < opcionais.length;
          g.groupSelected = marcados === opcionais.length;
        }
        var ga = comp.querySelector('[data-group-alert="' + gi + '"]');
        if (ga) ga.classList.toggle('d-none', marcados === opcionais.length);
        g.cookieList.forEach(function (c, ci) {
          var sw = comp.querySelector('#sw-cookie-' + gi + '-' + ci);
          if (sw) {
            sw.checked = c.cookieSelected;
            var lb = comp.querySelector('label[for="sw-cookie-' + gi + '-' + ci + '"]');
            if (lb) lb.textContent = c.cookieSelected ? 'Ligado' : 'Desligado';
          }
          var ca = comp.querySelector('[data-cookie-alert="' + gi + '-' + ci + '"]');
          if (ca) ca.classList.toggle('d-none', !!c.cookieSelected);
        });
      });
    }

    function jsonSaida() {
      return {
        selectAll: selecionados() === todosCookies().length ? true : (selecionados() === 0 ? false : 'indeterminated'),
        cookieGroups: DADOS.cookieGroups.map(function (g) {
          return {
            groupId: g.groupId,
            groupSelected: g.groupSelected,
            cookieList: g.cookieList.map(function (c) {
              return { cookieId: c.cookieId, cookieSelected: c.cookieSelected };
            })
          };
        })
      };
    }
    // Aplica de fato a escolha do usuário aos recursos opcionais
    function aplicar(saida) {
      var opc = {};
      saida.cookieGroups.forEach(function (g) {
        g.cookieList.forEach(function (c) { opc[c.cookieId] = c.cookieSelected; });
      });
      if (opc.vlibras === false) {
        d.querySelectorAll('[vw], .vlibras-plugin, #vlibras-wrapper').forEach(function (el) { el.remove(); });
      }
      try { localStorage.setItem(KEY, JSON.stringify(saida)); } catch (e) { /* não persiste entre sessões */ }
    }
    function restaurar() {
      var salvo;
      try { salvo = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { salvo = null; }
      if (!salvo) return false;
      var mapa = {};
      salvo.cookieGroups.forEach(function (g) {
        g.cookieList.forEach(function (c) { mapa[c.cookieId] = c.cookieSelected; });
      });
      DADOS.cookieGroups.forEach(function (g) {
        g.cookieList.forEach(function (c) {
          if (c.cookieOptOut && mapa[c.cookieId] === false) c.cookieSelected = false;
        });
      });
      aplicar(jsonSaida());
      return true;
    }

    function abrir(modo) {
      comp.classList.remove('d-none');
      if (modo === 'open') {
        comp.classList.remove('default');
        comp.querySelector('#cookiebarPolitics').classList.add('d-none');
        d.body.style.overflowY = 'hidden';
      }
      if (scrim) scrim.removeAttribute('hidden');
      comp.focus();
    }
    function recolher() {
      comp.classList.add('default');
      comp.querySelector('#cookiebarPolitics').classList.remove('d-none');
      d.body.style.overflowY = '';
    }
    function fechar() {
      comp.classList.add('d-none');
      if (scrim) scrim.setAttribute('hidden', '');
      d.body.style.overflowY = '';
    }

    construir();
    sincronizar();

    comp.addEventListener('change', function (ev) {
      var alvo = ev.target;
      if (alvo.id === 'ck-all') {
        todosCookies().forEach(function (c) { if (c.cookieOptOut) c.cookieSelected = alvo.checked; });
      } else if (alvo.hasAttribute('data-group')) {
        var g = DADOS.cookieGroups[Number(alvo.getAttribute('data-group'))];
        g.cookieList.forEach(function (c) { if (c.cookieOptOut) c.cookieSelected = alvo.checked; });
      } else if (alvo.hasAttribute('data-cookie')) {
        var pos = alvo.getAttribute('data-cookie').split('-');
        DADOS.cookieGroups[pos[0]].cookieList[pos[1]].cookieSelected = alvo.checked;
      } else return;
      sincronizar();
    });
    comp.addEventListener('click', function (ev) {
      var toggle = ev.target.closest('[data-toggle-group]');
      if (toggle) {
        var gi = toggle.getAttribute('data-toggle-group');
        var bloco = comp.querySelector('[data-group-block="' + gi + '"]');
        var aberto = bloco.classList.toggle('open');
        bloco.querySelectorAll('[data-toggle-group="' + gi + '"]').forEach(function (b) {
          b.setAttribute('aria-expanded', String(aberto));
        });
        var icone = bloco.querySelector('[data-toggle-group="' + gi + '"] i.fas');
        if (icone) { icone.classList.toggle('fa-angle-down', !aberto); icone.classList.toggle('fa-angle-up', aberto); }
        return;
      }
      if (ev.target.closest('#cookiebarPolitics')) { abrir('open'); return; }
      if (ev.target.closest('#cookiebarClose')) { recolher(); return; }
      if (ev.target.closest('#cookiebarAccept')) { aplicar(jsonSaida()); fechar(); }
    });
    d.addEventListener('click', function (ev) {
      if (ev.target.closest('#btnCookieSettings')) { abrir('open'); }
    });

    if (!restaurar()) abrir('default');
    else if (scrim) scrim.setAttribute('hidden', '');
  })();

  /* ── ONBOARDING — tour de apresentação em carrossel; abre sozinho no
     primeiro acesso (a menos que já dispensado) e pode ser reaberto por
     "Como usar este painel" (rodapé/menu) ou pelo switch dedicado. ── */
  (function () {
    var scrimEl = d.getElementById('onboardingScrim');
    if (!scrimEl) return;
    /* Controlador do componente Carousel (equivalente ao BRCarousel do
       core gov.br): palco com página ativa, botões de navegação
       desabilitados nos limites (navegação linear), indicador de
       páginas sincronizado, navegação por gesto (swipe) e por teclado,
       e reprodução automática opcional em looping. */
    var carousel = d.getElementById('onboardingCarousel');
    var col = d.getElementById('onboardingCol');
    var pages = Array.prototype.slice.call(col.querySelectorAll('.carousel-page'));
    var total = pages.length;
    var stepBtns = Array.prototype.slice.call(carousel.querySelectorAll('.step-progress-btn'));
    var stepText = d.getElementById('onboardingStepText');
    var prevBtn = carousel.querySelector('.carousel-btn-prev');
    var nextBtn = carousel.querySelector('.carousel-btn-next');
    var playBtn = carousel.querySelector('.carousel-btn-play');
    var pauseBtn = carousel.querySelector('.carousel-btn-pause');
    var startBtn = scrimEl.querySelector('.ob-start');
    var finishBtn = scrimEl.querySelector('.ob-finish');
    var skipLink = d.getElementById('onboardingSkip');
    var autoToggle = d.getElementById('obAutoToggle');
    var dontShow = d.getElementById('obDontShow');
    var current = 0;
    var timer = null;
    var DURACAO = 9000; // tempo de leitura estimado por página

    // Em reprodução automática a navegação é circular ("não pare na
    // última página"); na navegação manual é linear, com os botões
    // desabilitados na primeira e na última página.
    function isCircular() { return carousel.hasAttribute('data-circular') || !!timer; }

    function disabledBtns() {
      if (isCircular()) {
        prevBtn.removeAttribute('disabled');
        nextBtn.removeAttribute('disabled');
        return;
      }
      if (current === 0) {
        if (d.activeElement === prevBtn) nextBtn.focus();
        prevBtn.setAttribute('disabled', '');
      } else prevBtn.removeAttribute('disabled');
      if (current === total - 1) {
        if (d.activeElement === nextBtn) prevBtn.focus();
        nextBtn.setAttribute('disabled', '');
      } else nextBtn.removeAttribute('disabled');
    }
    function setActiveStage(num) {
      current = Math.max(0, Math.min(total - 1, num));
      pages.forEach(function (p, i) {
        p.removeAttribute('active');
        p.style.left = i > current ? '100%' : '-100%';
        if (i === current) { p.setAttribute('active', ''); p.style.left = '0'; }
      });
      stepBtns.forEach(function (b, i) {
        if (i === current) { b.setAttribute('active', ''); b.setAttribute('aria-selected', 'true'); }
        else { b.removeAttribute('active'); b.setAttribute('aria-selected', 'false'); }
      });
      if (stepText) stepText.textContent = 'Página ' + (current + 1) + ' de ' + total;
      disabledBtns();
    }
    function go(idx) { setActiveStage(idx); }
    function shiftPage(n) {
      var ultimo = total - 1;
      var alvo = current + (n < 0 ? -1 : 1);
      if (isCircular()) alvo = n < 0 ? (current === 0 ? ultimo : current - 1)
        : (current === ultimo ? 0 : current + 1);
      setActiveStage(Math.max(0, Math.min(ultimo, alvo)));
    }
    prevBtn.addEventListener('click', function () { shiftPage(-1); });
    nextBtn.addEventListener('click', function () { shiftPage(1); });
    stepBtns.forEach(function (b, i) { b.addEventListener('click', function () { pausar(); setActiveStage(i); }); });
    if (startBtn) startBtn.addEventListener('click', function () { go(1); });

    // Reprodução automática — começa sempre pausada e não é oferecida
    // em telas pequenas nem a quem pediu menos movimento.
    var semAuto = window.matchMedia && (window.matchMedia('(max-width: 991px)').matches ||
      window.matchMedia('(prefers-reduced-motion: reduce)').matches);
    if (semAuto && playBtn && pauseBtn) { playBtn.hidden = true; pauseBtn.hidden = true; }
    function reproduzir() {
      if (semAuto || timer) return;
      timer = setInterval(function () { shiftPage(1); }, DURACAO);
      if (playBtn) playBtn.hidden = true;
      if (pauseBtn) pauseBtn.hidden = false;
      if (pauseBtn) pauseBtn.focus();
      disabledBtns();
    }
    function pausar() {
      if (!timer) return;
      clearInterval(timer); timer = null;
      if (pauseBtn) pauseBtn.hidden = true;
      if (playBtn) playBtn.hidden = false;
      disabledBtns();
    }
    if (playBtn) playBtn.addEventListener('click', reproduzir);
    if (pauseBtn) pauseBtn.addEventListener('click', function () { pausar(); if (playBtn) playBtn.focus(); });

    // Navegação por gesto (swipe) em telas de toque
    var xIni = null;
    col.addEventListener('touchstart', function (ev) { xIni = ev.changedTouches[0].clientX; }, { passive: true });
    col.addEventListener('touchend', function (ev) {
      if (xIni === null) return;
      var dx = ev.changedTouches[0].clientX - xIni;
      xIni = null;
      if (Math.abs(dx) < 40) return;
      pausar();
      shiftPage(dx < 0 ? 1 : -1);
    }, { passive: true });

    // Navegação por teclado dentro do carrossel
    carousel.addEventListener('keydown', function (ev) {
      if (ev.key === 'ArrowRight') { pausar(); shiftPage(1); }
      else if (ev.key === 'ArrowLeft') { pausar(); shiftPage(-1); }
    });

    function fechar() { pausar(); scrimEl.hidden = true; }
    if (skipLink) skipLink.addEventListener('click', function (ev) { ev.preventDefault(); fechar(); });
    if (finishBtn) finishBtn.addEventListener('click', fechar);
    function openTour() { go(0); scrimEl.hidden = false; }

    function isDismissed() {
      try { return !!localStorage.getItem('painel_processos_onboarding_dismissed'); }
      catch (e) { return false; }
    }
    function setDismissed(value) {
      try {
        if (value) localStorage.setItem('painel_processos_onboarding_dismissed', '1');
        else localStorage.removeItem('painel_processos_onboarding_dismissed');
      } catch (e) { /* não persiste entre sessões */ }
    }
    if (autoToggle) {
      autoToggle.checked = !isDismissed();
      autoToggle.addEventListener('change', function () {
        setDismissed(!autoToggle.checked);
        if (dontShow) dontShow.checked = !autoToggle.checked;
      });
    }
    if (dontShow) {
      dontShow.checked = isDismissed();
      dontShow.addEventListener('change', function () {
        setDismissed(dontShow.checked);
        if (autoToggle) autoToggle.checked = !dontShow.checked;
      });
    }

    ['menuHelpTrigger', 'footerHelpTrigger'].forEach(function (id) {
      var btn = d.getElementById(id);
      if (!btn) return;
      btn.addEventListener('click', function (ev) {
        ev.preventDefault();
        var menu = d.getElementById('sectionMenu');
        if (menu && menu.classList.contains('active')) menu.classList.remove('active');
        openTour();
      });
    });

    if (!isDismissed()) setTimeout(openTour, 500);
  })();

  /* ── COMPARTILHAR / IMPRIMIR — presentes em toda página. Compartilhar usa
     a API nativa do navegador quando existe (menu de compartilhamento do
     sistema); sem suporte, copia o link pra área de transferência; sem
     isso também, mostra o link num prompt pra copiar manualmente. Nunca
     falha silenciosamente. ── */
  (function () {
    var btnShare = d.getElementById('btnCompartilhar');
    var btnPrint = d.getElementById('btnImprimir');
    if (btnPrint) btnPrint.addEventListener('click', function () { window.print(); });
    if (!btnShare) return;

    function tituloAtual() {
      var painelAtivo = d.querySelector('#mainTabContent .tab-panel:not([hidden])');
      var h = painelAtivo && painelAtivo.querySelector('h1, h2');
      return (h ? h.textContent.trim() + ' — ' : '') + 'Painel de Gestão de Processos · Codevasf';
    }
    function toast(msg) {
      var el = d.getElementById('ppShareToast');
      if (!el) {
        el = d.createElement('div');
        el.id = 'ppShareToast'; el.className = 'pp-share-toast'; el.setAttribute('role', 'status');
        d.body.appendChild(el);
      }
      el.textContent = msg; el.classList.add('show');
      clearTimeout(el._t); el._t = setTimeout(function () { el.classList.remove('show'); }, 2200);
    }
    btnShare.addEventListener('click', function () {
      var dados = { title: tituloAtual(), url: location.href };
      if (navigator.share) {
        navigator.share(dados).catch(function () { /* usuário cancelou o menu do sistema */ });
        return;
      }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(location.href).then(function () {
          toast('Link copiado para a área de transferência!');
        }).catch(function () { window.prompt('Copie o link para compartilhar:', location.href); });
        return;
      }
      window.prompt('Copie o link para compartilhar:', location.href);
    });
  })();

  /* ── Tooltips (Componente Tooltip, gov.br DS) ──────────────────────
     1) [data-tooltip-text] em elementos HTML normais (ex.: ícone de
        explicação nos gráficos do dashboard): cria/reaproveita um
        .tt-bubble filho e alterna .is-visible no hover/foco.
     2) Formas SVG dos gráficos (círculos, barras, arcos) já carregam
        um <title> nativo — aqui reaproveitamos esse mesmo texto para
        um balão flutuante (.graf-tip) com a identidade visual do DS,
        em vez do tooltip feio e lento do navegador. ── */
  (function () {
    document.addEventListener('mouseover', function (ev) {
      var el = ev.target.closest && ev.target.closest('[data-tooltip-text]');
      if (el) {
        var b = el.querySelector(':scope > .tt-bubble');
        if (!b) { b = document.createElement('span'); b.className = 'tt-bubble'; b.textContent = el.getAttribute('data-tooltip-text'); el.appendChild(b); }
        b.classList.add('is-visible');
      }
    });
    document.addEventListener('mouseout', function (ev) {
      var el = ev.target.closest && ev.target.closest('[data-tooltip-text]');
      if (el && !el.contains(ev.relatedTarget)) {
        var b = el.querySelector(':scope > .tt-bubble'); if (b) b.classList.remove('is-visible');
      }
    });
    document.addEventListener('focusin', function (ev) {
      var el = ev.target.closest && ev.target.closest('[data-tooltip-text]');
      if (el) {
        var b = el.querySelector(':scope > .tt-bubble');
        if (!b) { b = document.createElement('span'); b.className = 'tt-bubble'; b.textContent = el.getAttribute('data-tooltip-text'); el.appendChild(b); }
        b.classList.add('is-visible');
      }
    });
    document.addEventListener('focusout', function (ev) {
      var el = ev.target.closest && ev.target.closest('[data-tooltip-text]');
      if (el) { var b = el.querySelector(':scope > .tt-bubble'); if (b) b.classList.remove('is-visible'); }
    });

    var graftip = null;
    function graftipEl() {
      if (!graftip) { graftip = document.createElement('div'); graftip.className = 'graf-tip'; document.body.appendChild(graftip); }
      return graftip;
    }
    function posiciona(ev) {
      var t = graftipEl(), x = ev.clientX + 14, y = ev.clientY + 14;
      if (x + 260 > window.innerWidth) x = ev.clientX - 260 - 14;
      if (y + 60 > window.innerHeight) y = ev.clientY - 40;
      t.style.left = x + 'px'; t.style.top = y + 'px';
    }
    document.addEventListener('mouseover', function (ev) {
      var el = ev.target;
      if (el && el.namespaceURI === 'http://www.w3.org/2000/svg' && el.closest('.graf')) {
        var tt = el.querySelector(':scope > title');
        if (tt) { var t = graftipEl(); t.textContent = tt.textContent; t.classList.add('is-visible'); posiciona(ev); }
      }
    });
    document.addEventListener('mousemove', function (ev) {
      if (graftip && graftip.classList.contains('is-visible')) posiciona(ev);
    });
    document.addEventListener('mouseout', function (ev) {
      var el = ev.target;
      if (el && el.namespaceURI === 'http://www.w3.org/2000/svg' && el.querySelector(':scope > title') && graftip) {
        graftip.classList.remove('is-visible');
      }
    });
  })();

  /* ── NOTIFICATION — navegação por Tab dentro da central de alertas e
     botão fechar do cabeçalho (obrigatório na grid de 4 colunas). ── */
  d.addEventListener('click', function (ev) {
    var aba = ev.target.closest('#notifPanel .tab-nav button[data-panel]');
    if (aba) {
      var painel = aba.getAttribute('data-panel');
      $all('#notifPanel .tab-item').forEach(function (li) { li.classList.remove('active'); });
      $all('#notifPanel .tab-nav button').forEach(function (b) { b.setAttribute('aria-selected', 'false'); });
      aba.closest('.tab-item').classList.add('active');
      aba.setAttribute('aria-selected', 'true');
      $all('#notifPanel .tab-panel').forEach(function (p) {
        var ativo = p.id === painel;
        p.classList.toggle('active', ativo);
        p.hidden = !ativo;
      });
      return;
    }
    if (ev.target.closest('#notifPanel [data-dismiss="dropdown"]')) {
      var p2 = d.getElementById('notifPanel');
      var t = d.getElementById('btnNotifications');
      if (p2) p2.hidden = true;
      if (t) { t.setAttribute('aria-expanded', 'false'); t.focus(); }
    }
  });

  /* ── SELECT (BRSelect) — comportamento oficial: dropdown pelo campo ou
     pelo botão terciário (ícone angle-down/angle-up), filtro por
     autocomplete no próprio input, empty state "não encontramos" quando o
     filtro não retorna itens, seleção simples (radio, fecha a lista) ou
     múltipla (checkbox + Selecionar/Desselecionar todos), navegação por
     ArrowUp/ArrowDown/Enter/Space/Esc/Tab e aria completo. No múltiplo o
     campo mostra "primeiro item + (n)", como manda o comportamento de
     ajustes textuais. Ignora os selects da paginação, que têm
     controlador próprio. ── */
  window.BRSelectInit = function (raiz, aoMudar) {
    $all('.br-select[data-select]', raiz || d).forEach(function (comp) {
      var campo = comp.querySelector('.br-input input[type="text"]');
      var botao = comp.querySelector('[data-trigger]');
      var lista = comp.querySelector('.br-list');
      var chave = comp.getAttribute('data-select');
      var multi = comp.hasAttribute('multiple');
      var itens = $all('.br-item:not(.not-found)', lista);
      var todosItem = lista.querySelector('[data-all]');

      function entradas() { return itens.filter(function (i) { return !i.hasAttribute('data-all'); }); }
      function marcados() {
        return entradas().filter(function (i) { return i.querySelector('input').checked; });
      }
      function rotulo(i) { return i.querySelector('label').textContent; }
      function pintarCampo() {
        var m = marcados();
        if (!m.length) { campo.value = ''; campo.removeAttribute('title'); return; }
        campo.title = m.map(rotulo).join(', ');
        campo.value = m.length === 1 ? rotulo(m[0]) : rotulo(m[0]) + ' + (' + (m.length - 1) + ')';
      }
      function pintarTodos() {
        if (!todosItem) return;
        var alvo = entradas(), m = marcados().length;
        var chk = todosItem.querySelector('input');
        chk.checked = m > 0 && m === alvo.length;
        chk.indeterminate = m > 0 && m < alvo.length;
        todosItem.querySelector('label').textContent = chk.checked ? 'Desselecionar todos' : 'Selecionar todos';
        todosItem.classList.toggle('selected', chk.checked);
        todosItem.setAttribute('aria-selected', String(chk.checked));
      }
      function abrir(v) {
        if (v) lista.setAttribute('expanded', ''); else lista.removeAttribute('expanded');
        campo.setAttribute('aria-expanded', String(v));
        botao.setAttribute('aria-expanded', String(v));
        botao.setAttribute('aria-label', v ? 'Ocultar lista' : 'Exibir lista');
        var ic = botao.querySelector('i');
        ic.classList.toggle('fa-angle-up', v);
        ic.classList.toggle('fa-angle-down', !v);
        if (!v) { filtrar(''); pintarCampo(); }
      }
      function filtrar(txt) {
        var norm = function (s) { return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); };
        var q = norm(txt || ''), achou = false;
        entradas().forEach(function (i) {
          var ok = norm(rotulo(i)).indexOf(q) >= 0;
          i.classList.toggle('d-none', !ok);
          if (ok) achou = true;
        });
        if (todosItem) todosItem.classList.toggle('d-none', !!q);
        var vazio = lista.querySelector('.not-found');
        if (achou && vazio) vazio.remove();
        if (!achou && !vazio) {
          lista.insertAdjacentHTML('beforeend',
            '<div class="br-item not-found"><p><strong>Ops!</strong> Não encontramos o que você está procurando.</p></div>');
        }
      }
      function emitir() {
        comp.dispatchEvent(new CustomEvent('onChange', { bubbles: true, detail: comp }));
        if (aoMudar) aoMudar(chave, marcados().map(function (i) { return i.querySelector('input').value; }));
      }
      function visiveis() { return entradas().filter(function (i) { return !i.classList.contains('d-none'); }); }
      function mover(passo) {
        var vis = visiveis();
        if (!vis.length) return;
        var atual = vis.indexOf(d.activeElement.closest ? d.activeElement.closest('.br-item') : null);
        var prox = atual < 0 ? (passo > 0 ? 0 : vis.length - 1) : Math.min(vis.length - 1, Math.max(0, atual + passo));
        vis[prox].focus();
      }

      campo.addEventListener('click', function () { abrir(!lista.hasAttribute('expanded')); });
      botao.addEventListener('click', function () { abrir(!lista.hasAttribute('expanded')); });
      campo.addEventListener('input', function () { abrir(true); filtrar(this.value); });
      campo.addEventListener('keydown', function (ev) {
        if (ev.key === 'Escape') abrir(false);
        else if (ev.key === 'ArrowDown') { ev.preventDefault(); abrir(true); mover(1); }
        else if (ev.key === 'ArrowUp') { ev.preventDefault(); abrir(true); mover(-1); }
        else if (ev.key === 'Tab') abrir(false);
      });
      lista.addEventListener('keydown', function (ev) {
        var it = ev.target.closest('.br-item');
        if (ev.key === 'ArrowDown') { ev.preventDefault(); mover(1); }
        else if (ev.key === 'ArrowUp') { ev.preventDefault(); mover(-1); }
        else if (ev.key === 'Escape') { abrir(false); campo.focus(); }
        else if ((ev.key === 'Enter' || ev.key === ' ') && it) { ev.preventDefault(); it.querySelector('input').click(); }
        else if (ev.key === 'Tab') abrir(false);
      });

      entradas().forEach(function (item) {
        item.querySelector('input').addEventListener('change', function () {
          if (!multi) {
            entradas().forEach(function (o) {
              var on = o === item;
              o.classList.toggle('selected', on);
              o.setAttribute('aria-selected', String(on));
            });
            abrir(false);
          } else {
            item.classList.toggle('selected', this.checked);
            item.setAttribute('aria-selected', String(this.checked));
            pintarTodos();
          }
          pintarCampo();
          emitir();
        });
      });
      if (todosItem) {
        todosItem.querySelector('input').addEventListener('change', function () {
          var on = this.checked;
          entradas().forEach(function (o) {
            o.querySelector('input').checked = on;
            o.classList.toggle('selected', on);
            o.setAttribute('aria-selected', String(on));
          });
          pintarTodos();
          pintarCampo();
          emitir();
        });
      }
      // Fechamento por clique externo: exposto para o único ouvinte global
      // de documento (evitar um ouvinte por instância, que se acumularia a
      // cada redesenho da tela).
      comp.__fecharFora = function (alvo) { if (!comp.contains(alvo)) abrir(false); };
      pintarTodos();
      pintarCampo();
    });
  };

  /* ── SKIP LINK — a diretriz pede que a tecla Esc recolha o componente:
     tira o foco do item, o que devolve o item para fora da viewport. ── */
  d.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape') return;
    var item = ev.target.closest && ev.target.closest('.br-skiplink a');
    if (item) item.blur();
  });

  /* ── SELECT / PAGINATION — clique fora fecha as listas de opções e os
     dropdowns do botão reticências. ── */
  d.addEventListener('click', function (ev) {
    $all('.br-select[data-select]').forEach(function (c) {
      if (c.__fecharFora) c.__fecharFora(ev.target);
    });
    if (!ev.target.closest('.br-select')) {
      $all('.br-select [data-pag-select] .br-list[expanded], [data-pag-select] .br-list[expanded]').forEach(function (l) {
        l.removeAttribute('expanded');
        var c = l.parentElement.querySelector('input[type="text"]');
        if (c) c.setAttribute('aria-expanded', 'false');
      });
    }
    if (!ev.target.closest('.pagination-ellipsis')) {
      $all('.pagination-ellipsis .br-list').forEach(function (l) {
        if (!l.hidden) {
          l.hidden = true;
          var b = l.previousElementSibling;
          if (b) b.setAttribute('aria-expanded', 'false');
        }
      });
    }
  });

  /* ── MESSAGE (br-message) — botão fechar (Anatomia item 5): remove a
     mensagem da tela, como o BRAlert do core. Delegado, para valer
     também nas mensagens geradas pelas telas. ── */
  d.addEventListener('click', function (ev) {
    var fechar = ev.target.closest('.br-message .close');
    if (!fechar) return;
    var msg = fechar.closest('.br-message');
    if (msg && msg.parentNode) msg.parentNode.removeChild(msg);
  });

  /* ── LOADING (br-loading) — o tipo determinado recebe a porcentagem
     pelo atributo data-progress; o script a espelha em --loading-percent,
     que desenha o círculo dinâmico progressivo. Observa mudanças do
     atributo para acompanhar o progresso em tempo real. ── */
  (function () {
    function aplicar(el) {
      el.style.setProperty('--loading-percent', el.getAttribute('data-progress') || 0);
      if (!el.hasAttribute('aria-valuenow')) {
        el.setAttribute('aria-valuenow', el.getAttribute('data-progress') || 0);
      }
    }
    function varrer(escopo) {
      (escopo || d).querySelectorAll('.br-loading[data-progress]').forEach(aplicar);
    }
    varrer(d);
    new MutationObserver(function (muts) {
      muts.forEach(function (m) {
        if (m.type === 'attributes' && m.target.classList.contains('br-loading')) aplicar(m.target);
        else if (m.type === 'childNodes' || m.addedNodes) {
          m.addedNodes.forEach(function (n) { if (n.nodeType === 1) varrer(n); });
        }
      });
    }).observe(d.body, { subtree: true, childList: true, attributes: true, attributeFilter: ['data-progress'] });
  })();

  /* ── API mínima para o app preencher o menu de seções ────────────── */
  window.PPUI = {
    iniciarAccordions: iniciarAccordions,
    setMenuSections: function (itens) {
      var ul = $('#sectionMenuList');
      if (!ul) return;
      /* Componente List (vertical, densidade média): cada seção é um
         Item com as três áreas do padrão — suporte visual (ícone), área
         principal (nome da seção) e suporte complementar (quantidade de
         registros) — separados por divider e com a seção atual no
         estado selecionado. */
      var atual = location.hash || '#/';
      ul.innerHTML = itens.map(function (it, i) {
        var sel = (it.href === atual) || (it.href !== '#/' && atual.indexOf(it.href) === 0);
        return '<li>' +
          (i ? '<span class="br-divider" role="presentation"></span>' : '') +
          '<a class="br-item menu-item py-3' + (sel ? ' selected' : '') + '" role="listitem" href="' + it.href + '"' +
          (sel ? ' aria-current="page"' : '') + '>' +
          '<span class="content">' +
          '<span class="icon item-visual"><i class="fas ' + it.icone + '" aria-hidden="true"></i></span>' +
          '<span class="item-main">' + it.rotulo + '</span>' +
          (it.meta ? '<span class="item-support">' + it.meta + '</span>' : '') +
          '</span></a></li>';
      }).join('');
      ul.addEventListener('click', function () {
        var m = $('#sectionMenu'); if (m) m.classList.remove('active');
      });
    },
    fecharMenus: function () { closeAllDropdowns(null); }
  };
})();
