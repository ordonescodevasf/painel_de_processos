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
     item clicado e troca o ícone chevron-down/chevron-up. Com o atributo
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
        i.classList.toggle('fa-chevron-up', aberto);
        i.classList.toggle('fa-chevron-down', !aberto);
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

  /* ── CAROUSEL (br-carousel) — Componentes > Carousel, gov.br DS ──────
     Controlador genérico para qualquer .br-carousel[data-auto] (o tour de
     onboarding tem o próprio controlador dedicado, mais rico em regras de
     produto — este cobre qualquer outro carrossel de conteúdo do painel).
     Anatomia: palco + botões de navegação (obrigatórios); botões de
     reprodução e indicador de páginas — br-step — (opcionais, ambos
     implementados). data-circular ativa laço infinito; sem ele, os botões
     desabilitam nas pontas. Reprodução automática nunca começa sozinha
     (só por clique no botão Reproduzir) e não é oferecida em telas
     pequenas nem a quem pediu menos movimento — mesma regra do tour. ── */
  function BRCarousel(el) {
    this.el = el;
    this.pages = $all('.carousel-page', el);
    this.total = this.pages.length;
    if (!this.total) return;
    this.stepBtns = $all('.step-progress-btn', el);
    this.prevBtn = el.querySelector('.carousel-btn-prev');
    this.nextBtn = el.querySelector('.carousel-btn-next');
    this.playBtn = el.querySelector('.carousel-btn-play');
    this.pauseBtn = el.querySelector('.carousel-btn-pause');
    this.stage = el.querySelector('.carousel-stage');
    var ativa = this.pages.findIndex(function (p) { return p.hasAttribute('active'); });
    this.current = ativa >= 0 ? ativa : 0;
    this.timer = null;
    this.duracao = parseInt(el.getAttribute('data-duracao'), 10) || 7000;
    this._bind();
    this._disabledBtns();
  }
  BRCarousel.prototype._circular = function () { return this.el.hasAttribute('data-circular') || !!this.timer; };
  BRCarousel.prototype._render = function () {
    var self = this;
    this.pages.forEach(function (p, i) {
      if (i === self.current) p.setAttribute('active', ''); else p.removeAttribute('active');
    });
    this.stepBtns.forEach(function (b, i) {
      if (i === self.current) { b.setAttribute('active', ''); b.setAttribute('aria-selected', 'true'); }
      else { b.removeAttribute('active'); b.setAttribute('aria-selected', 'false'); }
    });
  };
  BRCarousel.prototype._disabledBtns = function () {
    if (this._circular()) {
      if (this.prevBtn) this.prevBtn.removeAttribute('disabled');
      if (this.nextBtn) this.nextBtn.removeAttribute('disabled');
      return;
    }
    if (this.prevBtn) {
      if (this.current === 0) {
        if (d.activeElement === this.prevBtn && this.nextBtn) this.nextBtn.focus();
        this.prevBtn.setAttribute('disabled', '');
      } else this.prevBtn.removeAttribute('disabled');
    }
    if (this.nextBtn) {
      if (this.current === this.total - 1) {
        if (d.activeElement === this.nextBtn && this.prevBtn) this.prevBtn.focus();
        this.nextBtn.setAttribute('disabled', '');
      } else this.nextBtn.removeAttribute('disabled');
    }
  };
  BRCarousel.prototype._goTo = function (idx) {
    this.current = Math.max(0, Math.min(this.total - 1, idx));
    this._render();
    this._disabledBtns();
  };
  BRCarousel.prototype._shift = function (n) {
    var last = this.total - 1, alvo;
    if (this._circular()) alvo = n < 0 ? (this.current === 0 ? last : this.current - 1) : (this.current === last ? 0 : this.current + 1);
    else alvo = this.current + (n < 0 ? -1 : 1);
    this._goTo(Math.max(0, Math.min(last, alvo)));
  };
  BRCarousel.prototype._pause = function () {
    if (!this.timer) return;
    clearInterval(this.timer); this.timer = null;
    if (this.pauseBtn) this.pauseBtn.hidden = true;
    if (this.playBtn) this.playBtn.hidden = false;
    this._disabledBtns();
  };
  BRCarousel.prototype._play = function () {
    var self = this;
    if (this._semAuto || this.timer || this.total < 2) return;
    this.timer = setInterval(function () { self._shift(1); }, this.duracao);
    if (this.playBtn) this.playBtn.hidden = true;
    if (this.pauseBtn) { this.pauseBtn.hidden = false; this.pauseBtn.focus(); }
    this._disabledBtns();
  };
  BRCarousel.prototype._bind = function () {
    var self = this;
    if (this.prevBtn) this.prevBtn.addEventListener('click', function () { self._pause(); self._shift(-1); });
    if (this.nextBtn) this.nextBtn.addEventListener('click', function () { self._pause(); self._shift(1); });
    this.stepBtns.forEach(function (b, i) { b.addEventListener('click', function () { self._pause(); self._goTo(i); }); });
    this._semAuto = window.matchMedia && (window.matchMedia('(max-width: 991px)').matches ||
      window.matchMedia('(prefers-reduced-motion: reduce)').matches);
    if (this._semAuto) { if (this.playBtn) this.playBtn.hidden = true; if (this.pauseBtn) this.pauseBtn.hidden = true; }
    if (this.playBtn) this.playBtn.addEventListener('click', function () { self._play(); });
    if (this.pauseBtn) this.pauseBtn.addEventListener('click', function () { self._pause(); if (self.playBtn) self.playBtn.focus(); });
    var xIni = null;
    if (this.stage) {
      this.stage.addEventListener('touchstart', function (ev) { xIni = ev.changedTouches[0].clientX; }, { passive: true });
      this.stage.addEventListener('touchend', function (ev) {
        if (xIni === null) return;
        var dx = ev.changedTouches[0].clientX - xIni; xIni = null;
        if (Math.abs(dx) < 40) return;
        self._pause(); self._shift(dx < 0 ? 1 : -1);
      }, { passive: true });
    }
    this.el.addEventListener('keydown', function (ev) {
      if (ev.key === 'ArrowRight') { self._pause(); self._shift(1); }
      else if (ev.key === 'ArrowLeft') { self._pause(); self._shift(-1); }
    });
  };
  function iniciarCarousels(escopo) {
    Array.prototype.forEach.call((escopo || d).querySelectorAll('.br-carousel[data-auto]'), function (el) {
      if (el.__brCarousel) return;
      el.__brCarousel = new BRCarousel(el);
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
    if (on) d.documentElement.setAttribute('data-contrast', 'enabled');
    else d.documentElement.removeAttribute('data-contrast');
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
      if (ic) { ic.classList.toggle('fa-chevron-up', abre); ic.classList.toggle('fa-chevron-down', !abre); }
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
      if (shotDataUrl) lines.push('', 'IMPORTANTE: anexe a este e-mail o print que foi baixado (botão "Baixar Print" no painel) antes de enviar.');
      var subject = encodeURIComponent('Erro no Repositório de Processos');
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
      optOutButton: 'Definir Cookies',
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
        '<button class="br-button circle small terciary" type="button" data-toggle-group="' + gi + '" aria-label="Expandir a classe ' + esc(g.groupName) + '" aria-expanded="false" aria-controls="cookie-list-' + gi + '"><i class="fas fa-chevron-down" aria-hidden="true"></i></button>' +
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
        /* role="list" explícito: o Componente List precisa expor a
           estrutura de lista ao leitor de tela — sem ele a div não é
           anunciada como lista nem informa quantos itens tem. */
        '<div class="br-list" role="list" aria-label="Classes de cookies">' +
        DADOS.cookieGroups.map(grupoHtml).join('') + '</div>' +
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
        if (icone) { icone.classList.toggle('fa-chevron-down', !aberto); icone.classList.toggle('fa-chevron-up', aberto); }
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
    var footNext = d.getElementById('onboardingNext');
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
      /* Botão primário do rodapé: "Iniciar" na abertura, "Avançar" no meio
         e "Concluir" na última página — mesma regra do Wizard. "Pular" sai
         de cena no fim, onde já não há tutorial a pular. */
      if (footNext) footNext.textContent = current === 0 ? 'Iniciar'
        : (current === total - 1 ? 'Concluir' : 'Próximo');
      if (skipLink) skipLink.hidden = current === total - 1;
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

    /* Sem a caixa "Não mostrar novamente" no rodapé, fechar o tutorial JÁ
       vale como visto — é o comportamento que a caixa dava, agora implícito
       (o tour continua acessível pelo menu de ajuda). */
    function fechar() { pausar(); scrimEl.hidden = true; setDismissed(true); }
    if (skipLink) skipLink.addEventListener('click', function (ev) { ev.preventDefault(); fechar(); });
    if (finishBtn) finishBtn.addEventListener('click', fechar);
    if (footNext) footNext.addEventListener('click', function () {
      pausar();
      if (current === total - 1) fechar(); else shiftPage(1);
    });
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
      return (h ? h.textContent.trim() + ' — ' : '') + 'Repositório de Processos · Codevasf';
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
     pelo botão terciário (ícone caret-down/caret-up), filtro por
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
        ic.classList.toggle('fa-caret-up', v);
        ic.classList.toggle('fa-caret-down', !v);
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

      function limpar() {
        entradas().forEach(function (o) {
          o.querySelector('input').checked = false;
          o.classList.remove('selected');
          o.setAttribute('aria-selected', 'false');
        });
        pintarTodos(); pintarCampo(); pintarLimpar(); emitir();
      }
      // Botão de limpar dentro do campo, visível só quando há escolha.
      var btnLimpar = d.createElement('button');
      btnLimpar.type = 'button';
      btnLimpar.className = 'br-button circle small select-limpar';
      btnLimpar.setAttribute('aria-label', 'Limpar seleção');
      btnLimpar.innerHTML = '<i class="fas fa-times" aria-hidden="true"></i>';
      btnLimpar.hidden = true;
      btnLimpar.addEventListener('click', function (ev) {
        ev.stopPropagation(); limpar(); campo.focus();
      });
      botao.parentNode.insertBefore(btnLimpar, botao);
      function pintarLimpar() { btnLimpar.hidden = !marcados().length; }

      entradas().forEach(function (item) {
        // Em seleção única o input é radio: um segundo clique no item já
        // escolhido não dispara change. O mousedown desmarca antes, e o
        // clique seguinte reemite a mudança — é o "desfazer" do controle.
        if (!multi) {
          item.addEventListener('mousedown', function () {
            var inp = item.querySelector('input');
            if (inp.checked) { setTimeout(limpar, 0); }
          });
        }
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
          pintarLimpar();
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
          pintarLimpar();
          emitir();
        });
      }
      // Fechamento por clique externo: exposto para o único ouvinte global
      // de documento (evitar um ouvinte por instância, que se acumularia a
      // cada redesenho da tela).
      comp.__fecharFora = function (alvo) { if (!comp.contains(alvo)) abrir(false); };
      pintarTodos();
      pintarCampo();
      pintarLimpar();
    });
  };

  /* ── TAB — navegação por teclado exigida pelo componente: setas
     esquerda/direita percorrem as abas, Home/End saltam para a primeira
     e a última. Enter e Space já são nativos do <button>. ── */
  d.addEventListener('keydown', function (ev) {
    var b = ev.target.closest && ev.target.closest('.tab-nav [role="tab"], .tab-nav .tab-item button');
    if (!b) return;
    var nav = b.closest('.tab-nav');
    var abas = $all('.tab-item button:not(:disabled)', nav);
    var i = abas.indexOf(b);
    if (i < 0) return;
    var alvo = -1;
    if (ev.key === 'ArrowRight') alvo = (i + 1) % abas.length;
    else if (ev.key === 'ArrowLeft') alvo = (i - 1 + abas.length) % abas.length;
    else if (ev.key === 'Home') alvo = 0;
    else if (ev.key === 'End') alvo = abas.length - 1;
    if (alvo < 0) return;
    ev.preventDefault();
    abas[alvo].focus();
  });

  /* ── TAG (br-tag) ── Tipo interação: o botão fechar dispensa a tag
     referenciada por data-dismiss (ou a própria tag que o contém). Tipo
     interação persistente: o estado .selected acompanha o input, com
     exclusividade entre radios de mesmo name. ── */
  d.addEventListener('click', function (ev) {
    var fechar = ev.target.closest && ev.target.closest('.br-tag [data-dismiss]');
    if (!fechar) return;
    var alvo = d.getElementById(fechar.getAttribute('data-dismiss')) || fechar.closest('.br-tag');
    if (alvo) alvo.remove();
  });
  d.addEventListener('change', function (ev) {
    var input = ev.target;
    if (!input.matches || !input.matches('.br-tag.interaction-select input')) return;
    if (input.type === 'radio' && input.name) {
      $all('.br-tag.interaction-select input[name="' + input.name + '"]').forEach(function (o) {
        o.closest('.br-tag').classList.toggle('selected', o.checked);
      });
    } else {
      input.closest('.br-tag').classList.toggle('selected', input.checked);
    }
  });

  /* ── TEXTAREA (br-textarea) ── Contador de caracteres: com limite
     (maxlength) o texto passa de "limite máximo" para "restam N"; sem
     limite, conta os caracteres digitados. Ambos em região aria-live,
     como no comportamento oficial. ── */
  function atualizarContador(campo) {
    var caixa = campo.closest('.br-textarea');
    if (!caixa) return;
    var limite = caixa.querySelector('.limit');
    var atual = caixa.querySelector('.current');
    var chars = caixa.querySelector('.characters');
    var n = campo.value.length;
    var max = campo.getAttribute('maxlength');
    if (chars) { chars.innerHTML = '<strong>' + n + '</strong> caracteres digitados'; return; }
    if (!max) return;
    if (n === 0) {
      if (limite) limite.innerHTML = 'Limite máximo de <strong>' + max + '</strong> caracteres';
      if (atual) atual.innerHTML = '';
    } else {
      if (limite) limite.innerHTML = '';
      if (atual) atual.innerHTML = 'Restam <strong>' + (max - n) + '</strong> caracteres';
    }
  }
  ['input', 'focus'].forEach(function (evt) {
    d.addEventListener(evt, function (ev) {
      if (ev.target.matches && ev.target.matches('.br-textarea textarea')) atualizarContador(ev.target);
    }, true);
  });

  /* ── TABLE (br-table) ── Comportamento genérico das tabelas montadas por
     tabelaGov: troca de densidade pelo menu de ações, busca que cobre a
     barra de título e filtra as linhas, e ordenação por coluna com
     aria-sort + ícone (sort / sort-up / sort-down). Delegado no documento,
     então vale para tabelas renderizadas depois. ── */
  function ordenarTabela(th, tabela) {
    var head = th.closest('tr');
    var idx = Array.prototype.indexOf.call(head.children, th);
    var atual = th.getAttribute('aria-sort');
    // Ciclo de três estados: crescente → decrescente → sem ordenação. O
    // componente prevê os três (seta dupla = ordenação padrão), e sem o
    // terceiro o usuário não tem como desfazer o que ordenou.
    var dir = atual === 'ascending' ? 'descending' : (atual === 'descending' ? '' : 'ascending');
    var corpo = tabela.querySelector('tbody');
    // A ordem original é a que a tela montou: guardada na primeira
    // ordenação, é para onde o terceiro clique volta.
    if (!corpo._ordemOriginal) corpo._ordemOriginal = $all('tr', corpo);
    $all('th', head).forEach(function (o) {
      o.removeAttribute('aria-sort');
      var i = o.querySelector('.sort-btn i');
      if (i) i.className = 'fas fa-sort';
      o.classList.remove('is-sorted');
    });
    var ic = th.querySelector('.sort-btn i');
    if (!dir) {
      corpo._ordemOriginal.forEach(function (tr) { corpo.appendChild(tr); });
      return;
    }
    th.setAttribute('aria-sort', dir);
    th.classList.add('is-sorted');
    if (ic) ic.className = 'fas ' + (dir === 'ascending' ? 'fa-sort-up' : 'fa-sort-down');
    var linhas = $all('tr', corpo).filter(function (tr) { return !tr.classList.contains('collapse'); });
    var sinal = dir === 'ascending' ? 1 : -1;
    linhas.sort(function (a, b) {
      var va = (a.children[idx] || {}).textContent || '';
      var vb = (b.children[idx] || {}).textContent || '';
      return sinal * va.trim().localeCompare(vb.trim(), 'pt-BR', { numeric: true, sensitivity: 'base' });
    });
    linhas.forEach(function (tr) {
      // A linha de detalhe é irmã imediata da sua linha de dados e viaja
      // junto: reanexada logo depois, o par não se desfaz na ordenação.
      var det = tr.nextElementSibling;
      corpo.appendChild(tr);
      if (det && det.classList.contains('collapse')) corpo.appendChild(det);
    });
  }
  /* ── TABLE · Seleção de linhas e Barra Contextual ───────────────────
     Comportamento 4 da anatomia do Componente Table: caixa de seleção no
     início de cada linha, caixa "selecionar tudo" no header (com o estado
     intermediário do utilitário Checkgroup) e Barra Contextual que surge
     sob a Barra de Título com a contagem e as ações da seleção.

     A coluna de seleção é injetada aqui, em vez de escrita em cada tabela:
     as linhas são montadas por dezenas de trechos diferentes em app.js, e
     acrescentar uma célula em cada um deles multiplicaria o mesmo código.
     O CSV também sai do DOM — o texto visível de cada célula —, então
     qualquer tabela nova ganha a exportação sem escrever nada. ── */
  function tblMontarCollapse(t) {
    var tabela = t.querySelector('table');
    if (!tabela || t._collapsePronto) return;
    /* Tabela que já monta a própria coluna de expandir (Documentos
       publicados) fica de fora: a injeção acrescentaria uma segunda
       célula em cada linha e o corpo passaria a ter uma coluna a mais
       que o cabeçalho — todo valor sob o rótulo errado. Mesmo cuidado
       que tblMontarSelecao tem com a coluna de seleção própria. */
    if (tabela.querySelector('thead .column-collapse')) { t._collapsePronto = true; return; }
    var detalhes = $all('tbody > tr.collapse', tabela);
    if (!detalhes.length) return;
    t._collapsePronto = true;
    t.setAttribute('data-collapse', 'data-collapse');
    var seq = 0, base = (t.id || 'tbl') + '-det';
    detalhes.forEach(function (tr) {
      var linha = tr.previousElementSibling;
      if (!linha || linha.classList.contains('collapse')) return;
      var alvo = tr.querySelector('td');
      if (!alvo) return;
      var id = base + '-' + (++seq);
      alvo.id = id;
      alvo.setAttribute('aria-hidden', 'true');
      alvo.hidden = true;
      var td = d.createElement('td');
      td.className = 'column-collapse';
      td.innerHTML = '<button class="br-button circle small" type="button" data-toggle="collapse"' +
        ' data-target="' + id + '" aria-controls="' + id + '" aria-expanded="false"' +
        ' aria-label="Expandir ou retrair os detalhes desta linha">' +
        '<i class="fas fa-chevron-down" aria-hidden="true"></i></button>';
      linha.insertBefore(td, linha.firstChild);
    });
    // Cabeçalho e colgroup acompanham a coluna nova.
    var thead = tabela.querySelector('thead tr');
    if (thead && !thead.querySelector('.column-collapse')) {
      var th = d.createElement('td');
      th.className = 'column-collapse';
      th.setAttribute('aria-hidden', 'true');
      thead.insertBefore(th, thead.firstChild);
    }
    var cg = tabela.querySelector('colgroup');
    if (cg && !cg.querySelector('col.col-collapse')) {
      var col = d.createElement('col');
      col.className = 'col-collapse';
      cg.insertBefore(col, cg.firstChild);
    }
    var total = thead ? thead.children.length : 0;
    detalhes.forEach(function (tr) {
      var td = tr.querySelector('td');
      if (td) td.setAttribute('colspan', String(total));
    });
  }
  function tblLinhasDados(t) {
    return $all('tbody > tr', t).filter(function (tr) {
      if (tr.classList.contains('collapse')) return false;
      return !tr.classList.contains('tbl-sem-resultado') && !tr.classList.contains('collapse');
    });
  }
  function tblMontarSelecao(t) {
    if (t._selPronta) return;
    t._selPronta = true;
    // Tabela que já traz coluna de seleção própria fica com ela.
    if (t.querySelector('thead .column-checkbox')) return;
    var id = t.id || ('tbl' + Math.random().toString(36).slice(2, 7));
    var cg = t.querySelector('colgroup');
    if (cg) cg.insertAdjacentHTML('afterbegin', '<col class="col-checkbox">');
    var trHead = t.querySelector('thead tr');
    if (trHead) {
      trHead.insertAdjacentHTML('afterbegin',
        '<th class="column-checkbox" scope="col">' +
        '<div class="br-checkbox hidden-label">' +
        '<input id="' + id + '-all" name="' + id + '-all" type="checkbox"' +
        ' data-parent="' + id + '-grp" aria-label="Selecionar todas as linhas">' +
        '<label for="' + id + '-all">Selecionar todas as linhas</label>' +
        '</div></th>');
    }
    tblLinhasDados(t).forEach(function (tr, i) {
      tr.insertAdjacentHTML('afterbegin',
        '<td class="column-checkbox">' +
        '<div class="br-checkbox hidden-label">' +
        '<input id="' + id + '-l' + i + '" name="' + id + '-l' + i + '" type="checkbox"' +
        ' data-child="' + id + '-grp" aria-label="Selecionar linha ' + (i + 1) + '">' +
        '<label for="' + id + '-l' + i + '">Selecionar linha ' + (i + 1) + '</label>' +
        '</div></td>');
    });
    // A linha de conteúdo expandido atravessa a tabela inteira: com a
    // coluna nova, o colspan precisa crescer junto.
    $all('tbody > tr.collapse > td[colspan]', t).forEach(function (td) {
      td.setAttribute('colspan', String(+td.getAttribute('colspan') + 1));
    });
    var cab = t.querySelector('.table-header');
    if (cab && !cab.querySelector('.selected-bar')) {
      cab.insertAdjacentHTML('beforeend',
        '<div class="selected-bar" role="status">' +
        '<div class="info"><span class="count">0</span>' +
        '<span class="text">item selecionado</span></div>' +
        '<button class="br-button circle small inverted" type="button" data-sel-acao="csv"' +
        ' aria-label="Exportar seleção em CSV"><i class="fas fa-file-csv" aria-hidden="true"></i></button>' +
        '<button class="br-button circle small inverted" type="button" data-sel-acao="limpar"' +
        ' aria-label="Limpar seleção"><i class="fas fa-times" aria-hidden="true"></i></button>' +
        '</div>');
    }
  }
  function tblAtualizarSelecao(t) {
    var linhas = tblLinhasDados(t);
    var marcadas = linhas.filter(function (tr) {
      var c = tr.querySelector('.column-checkbox input');
      return c && c.checked;
    });
    linhas.forEach(function (tr) {
      var c = tr.querySelector('.column-checkbox input');
      tr.classList.toggle('is-selected', !!(c && c.checked));
    });
    var pai = t.querySelector('thead .column-checkbox input');
    if (pai) {
      // checked + indeterminate juntos no estado intermediário, como o
      // utilitário Checkgroup define.
      pai.checked = marcadas.length > 0;
      pai.indeterminate = marcadas.length > 0 && marcadas.length < linhas.length;
    }
    var barra = t.querySelector('.selected-bar');
    if (!barra) return;
    barra.classList.toggle('show', marcadas.length > 0);
    barra.querySelector('.count').textContent = marcadas.length;
    barra.querySelector('.text').textContent =
      marcadas.length === 1 ? 'item selecionado' : 'itens selecionados';
  }
  function tblCsvSelecao(t) {
    var cabs = $all('thead th', t).slice(1).map(function (th) {
      return (th.textContent || '').trim();
    });
    var linhas = tblLinhasDados(t).filter(function (tr) {
      var c = tr.querySelector('.column-checkbox input');
      return c && c.checked;
    }).map(function (tr) {
      return $all('td', tr).slice(1).map(function (td) {
        var blocos = $all(':scope > div, :scope > p, :scope > ul', td);
        if (!blocos.length) return (td.textContent || '').replace(/\s+/g, ' ').trim();
        var clone = td.cloneNode(true);
        $all(':scope > div, :scope > p, :scope > ul', clone).forEach(function (b) { b.remove(); });
        return [clone.textContent].concat(blocos.map(function (b) { return b.textContent; }))
          .map(function (x) { return (x || '').replace(/\s+/g, ' ').trim(); })
          .filter(Boolean).join(' · ');
      });
    });
    var nome = (t.querySelector('.table-title') || {}).textContent || 'tabela';
    var matriz = [cabs].concat(linhas);
    var txt = matriz.map(function (l) {
      return l.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(';');
    }).join('\r\n');
    // BOM + ponto e vírgula: abre direto no Excel em pt-BR.
    var blob = new Blob(['\ufeff' + txt], { type: 'text/csv;charset=utf-8' });
    var a = d.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nome.trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') + '.csv';
    d.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  window.BRTableInit = function (raiz) {
    /* A coluna de expandir é a PRIMEIRA da linha e a de seleção vem logo
       depois (anatomia do componente). As duas inserem no início da
       linha, então quem roda por último é que fica em primeiro: a
       seleção entra antes, o collapse depois. */
    $all('.br-table[data-selection]', raiz || d).forEach(tblMontarSelecao);
    $all('.br-table', raiz || d).forEach(tblMontarCollapse);
  };
  d.addEventListener('change', function (ev) {
    var cx = ev.target.closest && ev.target.closest('.br-table[data-selection] .column-checkbox input');
    if (!cx) return;
    var t = cx.closest('.br-table');
    if (cx.closest('thead')) {
      /* A decisão vem da CONTAGEM anterior, não do estado do checkbox
         depois do clique: numa caixa marcada + indeterminada o navegador
         zera os dois antes do change disparar, e ler dali faria o clique
         a partir do estado parcial desselecionar tudo. Pela anatomia do
         Table, clicar no header a partir do parcial seleciona todas. */
      var linhas = tblLinhasDados(t);
      var marcadas = linhas.filter(function (tr) {
        var c = tr.querySelector('.column-checkbox input');
        return c && c.checked;
      }).length;
      var marcar = marcadas < linhas.length;
      linhas.forEach(function (tr) {
        var c = tr.querySelector('.column-checkbox input');
        if (c) c.checked = marcar;
      });
    }
    tblAtualizarSelecao(t);
  });
  d.addEventListener('click', function (ev) {
    var b = ev.target.closest && ev.target.closest('[data-sel-acao]');
    if (!b) return;
    var t = b.closest('.br-table');
    if (b.getAttribute('data-sel-acao') === 'csv') { tblCsvSelecao(t); return; }
    tblLinhasDados(t).forEach(function (tr) {
      var c = tr.querySelector('.column-checkbox input');
      if (c) c.checked = false;
    });
    tblAtualizarSelecao(t);
  });

  d.addEventListener('click', function (ev) {
    var raiz = ev.target.closest && ev.target.closest('.br-table[data-generic]');
    if (!raiz) return;
    var dens = ev.target.closest('[data-density]');
    if (dens) {
      ['small', 'medium', 'large'].forEach(function (c) { raiz.classList.remove(c); });
      raiz.classList.add(dens.getAttribute('data-density'));
      $all('[data-density]', raiz).forEach(function (o) { o.removeAttribute('aria-current'); });
      dens.setAttribute('aria-current', 'true');
      var menu = dens.closest('.br-list');
      if (menu) {
        menu.hidden = true;
        var t = raiz.querySelector('[data-target="' + menu.id + '"]');
        if (t) t.setAttribute('aria-expanded', 'false');
      }
      return;
    }
    var abrir = ev.target.closest('[data-toggle="search"]');
    if (abrir) {
      var barra = raiz.querySelector('.search-bar');
      barra.classList.add('show');
      abrir.setAttribute('aria-expanded', 'true');
      var campo = barra.querySelector('input');
      if (campo) campo.focus();
      return;
    }
    if (ev.target.closest('[data-dismiss="search"]')) {
      var b2 = raiz.querySelector('.search-bar');
      var c2 = b2.querySelector('input');
      b2.classList.remove('show');
      if (c2) { c2.value = ''; c2.dispatchEvent(new Event('input', { bubbles: true })); }
      var g2 = raiz.querySelector('[data-toggle="search"]');
      if (g2) { g2.setAttribute('aria-expanded', 'false'); g2.focus(); }
      return;
    }
    var sort = ev.target.closest('.sort-btn');
    if (sort) ordenarTabela(sort.closest('th'), raiz);
  });
  d.addEventListener('input', function (ev) {
    var raiz = ev.target.closest && ev.target.closest('.br-table[data-generic] .search-bar input');
    if (!raiz) return;
    var tabela = ev.target.closest('.br-table');
    var q = ev.target.value.trim().toLowerCase();
    var achou = 0;
    $all('tbody tr:not(.tbl-sem-resultado)', tabela).forEach(function (tr) {
      var ok = !q || tr.textContent.toLowerCase().indexOf(q) >= 0;
      tr.hidden = !ok;
      if (ok) achou++;
    });
    var vazio = tabela.querySelector('.tbl-sem-resultado');
    if (!achou && !vazio) {
      var cols = tabela.querySelectorAll('thead th').length;
      tabela.querySelector('tbody').insertAdjacentHTML('beforeend',
        '<tr class="tbl-sem-resultado"><td colspan="' + cols + '">Nenhum registro encontrado para a busca.</td></tr>');
    } else if (achou && vazio) vazio.remove();
    var total = tabela.querySelector('.tbl-total strong');
    if (total) total.textContent = achou;
  });

  /* ── TOOLTIP (br-tooltip) ── Tipo padrão: aparece no mouse over e no foco
     do gatilho, some no mouse out, no blur ou com Esc. Tipo popover: abre
     no clique (ou já vem com [active]) e só fecha no botão fechar, no
     clique externo ou por [timer]. A seta e o botão fechar do popover são
     criados aqui, como no comportamento oficial. ── */
  function ttGatilho(tt) {
    var pai = tt.parentElement;
    if (!pai) return null;
    // O gatilho pode ser o próprio contêiner (caso dos marcos do mapeamento,
    // em que o botão da etapa vem desabilitado e não dispara eventos de
    // mouse — quem escuta o hover é o contêiner em volta).
    if (pai.hasAttribute('data-tooltip-trigger')) return pai;
    return pai.querySelector('[data-tooltip-trigger]') || tt.previousElementSibling;
  }
  function ttMostrar(tt) {
    var timer = tt.getAttribute('timer');
    ttCancelarSaida(tt);
    tt.setAttribute('data-show', '');
    /* Reposicionamento: o balão nasce centralizado no gatilho e, perto das
       bordas da tela, escaparia da área visível — nas trilhas de marcos, a
       primeira etapa fica colada na margem esquerda. Desloca no eixo X o
       necessário para caber, mantendo a seta apontando para o gatilho. */
    var lugar = tt.getAttribute('place') || 'top';
    if (lugar === 'top' || lugar === 'bottom') {
      tt.style.setProperty('--tt-shift', '0px');
      var cx = tt.getBoundingClientRect(), folga = 8, dx = 0;
      if (cx.left < folga) dx = folga - cx.left;
      else if (cx.right > window.innerWidth - folga) dx = (window.innerWidth - folga) - cx.right;
      if (dx) tt.style.setProperty('--tt-shift', Math.round(dx) + 'px');
    }
    if (tt._t) clearTimeout(tt._t);
    if (timer) tt._t = setTimeout(function () { tt.removeAttribute('data-show'); }, +timer);
  }
  function ttEsconder(tt) { tt.removeAttribute('data-show'); if (tt._t) clearTimeout(tt._t); }
  /* WCAG 2.2 · 1.4.13 (Conteúdo ao passar o mouse ou receber foco) e a
     diretriz de acessibilidade do Tooltip: o balão não pode sumir no
     instante em que o cursor sai — o usuário precisa de tempo para ler e
     de poder levar o cursor até o próprio balão (para copiar o texto, por
     exemplo). Daí a carência antes de esconder, cancelada se o cursor
     voltar ao gatilho ou entrar no balão. */
  function ttAgendarSaida(tt) {
    if (tt._saida) clearTimeout(tt._saida);
    tt._saida = setTimeout(function () { ttEsconder(tt); }, 400);
  }
  function ttCancelarSaida(tt) { if (tt._saida) { clearTimeout(tt._saida); tt._saida = null; } }
  /* ── MESSAGE (br-message) ── Diretriz de acessibilidade do componente:
     o leitor de tela precisa anunciar QUAL é o tipo da mensagem (erro,
     aviso, sucesso ou informação) — WCAG 4.1.2. O ícone que carrega essa
     informação é aria-hidden, então sem um rótulo textual o leitor lê só
     o corpo da mensagem. Este init insere o rótulo invisível uma vez por
     mensagem; role="alert"/"status" já cobrem o anúncio dinâmico. ── */
  var MSG_ROTULO = { danger: 'Erro', warning: 'Aviso', success: 'Sucesso', info: 'Informação' };
  /* ── LOADING (br-loading) ── Arte OFICIAL do componente: os SVGs de
     img/loading/ vêm do pacote do DS e já são desenhados para receber a
     cor por custom property, então o mesmo desenho serve ao tema claro e
     ao contraste — os fills usam nomes neutros (--br-loading-color-*) e
     quem troca o tema é o CSS. O caminho é embutido
     aqui, e não buscado por fetch, porque o painel também roda aberto
     direto do arquivo (file://), onde fetch é bloqueado.
     Simplificação: o filtro de sombra interna do Figma foi descartado —
     seus ids colidiriam entre múltiplos Loadings na mesma página e o
     efeito é imperceptível no tamanho em que o componente roda. ── */
  var LOADING_SVG = {
    "small-main": {
      "vb": "0 0 19 19",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M9.5 0C14.7467 0 19 4.25329 19 9.5C19 14.7467 14.7467 19 9.5 19C4.25329 19 0 14.7467 0 9.5C0 4.25329 4.25329 0 9.5 0ZM9.5 4C6.46243 4 4 6.46243 4 9.5C4 12.5376 6.46243 15 9.5 15C12.5376 15 15 12.5376 15 9.5C15 6.46243 12.5376 4 9.5 4Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 19 19)\" d=\"M17.005 9.5C18.1068 9.5 19.021 10.4028 18.7913 11.4804C18.6709 12.0453 18.4989 12.5993 18.2769 13.1355C17.7994 14.2881 17.0997 15.3354 16.2175 16.2175C15.3354 17.0997 14.2881 17.7994 13.1355 18.2769C12.5993 18.4989 12.0453 18.6709 11.4804 18.7913C10.4028 19.021 9.5 18.1068 9.5 17.005V17.005C9.5 15.9032 10.4213 15.043 11.4517 14.6528C11.5043 14.6328 11.5566 14.6121 11.6086 14.5906C12.2771 14.3137 12.8845 13.9078 13.3962 13.3962C13.9078 12.8845 14.3137 12.2771 14.5906 11.6086C14.6121 11.5566 14.6328 11.5043 14.6528 11.4517C15.043 10.4213 15.9032 9.5 17.005 9.5V9.5Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    },
    "small-contrast": {
      "vb": "0 0 19 19",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M9.5 0C14.7467 0 19 4.25329 19 9.5C19 14.7467 14.7467 19 9.5 19C4.25329 19 0 14.7467 0 9.5C0 4.25329 4.25329 0 9.5 0ZM9.5 4C6.46243 4 4 6.46243 4 9.5C4 12.5376 6.46243 15 9.5 15C12.5376 15 15 12.5376 15 9.5C15 6.46243 12.5376 4 9.5 4Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 19 19)\" d=\"M17.005 9.5C18.1068 9.5 19.021 10.4028 18.7913 11.4804C18.6709 12.0453 18.4989 12.5993 18.2769 13.1355C17.7994 14.2881 17.0997 15.3354 16.2175 16.2175C15.3354 17.0997 14.2881 17.7994 13.1355 18.2769C12.5993 18.4989 12.0453 18.6709 11.4804 18.7913C10.4028 19.021 9.5 18.1068 9.5 17.005V17.005C9.5 15.9032 10.4213 15.043 11.4517 14.6528C11.5043 14.6328 11.5566 14.6121 11.6086 14.5906C12.2771 14.3137 12.8845 13.9078 13.3962 13.3962C13.9078 12.8845 14.3137 12.2771 14.5906 11.6086C14.6121 11.5566 14.6328 11.5043 14.6528 11.4517C15.043 10.4213 15.9032 9.5 17.005 9.5V9.5Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    },
    "medium-main": {
      "vb": "0 0 40 40",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M20 0C31.0457 0 40 8.9543 40 20C40 31.0457 31.0457 40 20 40C8.9543 40 0 31.0457 0 20C0 8.9543 8.9543 0 20 0ZM20 6C12.268 6 6 12.268 6 20C6 27.732 12.268 34 20 34C27.732 34 34 27.732 34 20C34 12.268 27.732 6 20 6Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 40 40)\" d=\"M37 20C38.6569 20 40.023 21.3505 39.7754 22.9888C39.5339 24.5869 39.0989 26.1536 38.4776 27.6537C37.4725 30.0802 35.9993 32.285 34.1421 34.1421C32.285 35.9993 30.0802 37.4725 27.6537 38.4776C26.1536 39.0989 24.5869 39.5339 22.9888 39.7754C21.3505 40.023 20 38.6569 20 37V37C20 35.3431 21.3582 34.0321 22.9771 33.6798C23.7893 33.503 24.586 33.2539 25.3576 32.9343C27.0561 32.2307 28.5995 31.1995 29.8995 29.8995C31.1995 28.5995 32.2307 27.0561 32.9343 25.3576C33.2539 24.586 33.503 23.7893 33.6798 22.9771C34.0321 21.3582 35.3431 20 37 20V20Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    },
    "medium-contrast": {
      "vb": "0 0 40 40",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M20 0C31.0457 0 40 8.9543 40 20C40 31.0457 31.0457 40 20 40C8.9543 40 0 31.0457 0 20C0 8.9543 8.9543 0 20 0ZM20 6C12.268 6 6 12.268 6 20C6 27.732 12.268 34 20 34C27.732 34 34 27.732 34 20C34 12.268 27.732 6 20 6Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 40 40)\" d=\"M37 20C38.6569 20 40.023 21.3505 39.7754 22.9888C39.5339 24.5869 39.0989 26.1536 38.4776 27.6537C37.4725 30.0802 35.9993 32.285 34.1421 34.1421C32.285 35.9993 30.0802 37.4725 27.6537 38.4776C26.1536 39.0989 24.5869 39.5339 22.9888 39.7754C21.3505 40.023 20 38.6569 20 37V37C20 35.3431 21.3582 34.0321 22.9771 33.6798C23.7893 33.503 24.586 33.2539 25.3576 32.9343C27.0561 32.2307 28.5995 31.1995 29.8995 29.8995C31.1995 28.5995 32.2307 27.0561 32.9343 25.3576C33.2539 24.586 33.503 23.7893 33.6798 22.9771C34.0321 21.3582 35.3431 20 37 20V20Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    },
    "large-main": {
      "vb": "0 0 72 72",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M36 0C55.8823 0 72 16.1177 72 36C72 55.8823 55.8823 72 36 72C16.1177 72 0 55.8823 0 36C0 16.1177 16.1177 0 36 0ZM36 8C20.536 8 8 20.536 8 36C8 51.464 20.536 64 36 64C51.464 64 64 51.464 64 36C64 20.536 51.464 8 36 8Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 72 72)\" d=\"M68.04 36C70.227 36 72.0225 37.7782 71.7824 39.952C71.4101 43.323 70.5626 46.6311 69.2597 49.7766C67.4505 54.1443 64.7988 58.1129 61.4558 61.4558C58.1129 64.7988 54.1443 67.4505 49.7766 69.2597C46.6311 70.5626 43.323 71.4101 39.952 71.7824C37.7782 72.0225 36 70.227 36 68.04V68.04C36 65.853 37.7816 64.1086 39.9469 63.8012C42.2771 63.4704 44.5617 62.8472 46.7457 61.9425C50.1526 60.5314 53.2481 58.463 55.8556 55.8556C58.463 53.2481 60.5314 50.1526 61.9425 46.7458C62.8472 44.5617 63.4704 42.2771 63.8012 39.9469C64.1086 37.7816 65.853 36 68.04 36V36Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    },
    "large-contrast": {
      "vb": "0 0 72 72",
      "d": "<path fill-rule=\"evenodd\" clip-rule=\"evenodd\" d=\"M36 0C55.8823 0 72 16.1177 72 36C72 55.8823 55.8823 72 36 72C16.1177 72 0 55.8823 0 36C0 16.1177 16.1177 0 36 0ZM36 8C20.536 8 8 20.536 8 36C8 51.464 20.536 64 36 64C51.464 64 64 51.464 64 36C64 20.536 51.464 8 36 8Z\" id=\"background_bg_0\" fill=\"var(--br-loading-color-background)\"/><path id=\"stroke\" transform=\"matrix(-1 0 0 -1 72 72)\" d=\"M68.04 36C70.227 36 72.0225 37.7782 71.7824 39.952C71.4101 43.323 70.5626 46.6311 69.2597 49.7766C67.4505 54.1443 64.7988 58.1129 61.4558 61.4558C58.1129 64.7988 54.1443 67.4505 49.7766 69.2597C46.6311 70.5626 43.323 71.4101 39.952 71.7824C37.7782 72.0225 36 70.227 36 68.04V68.04C36 65.853 37.7816 64.1086 39.9469 63.8012C42.2771 63.4704 44.5617 62.8472 46.7457 61.9425C50.1526 60.5314 53.2481 58.463 55.8556 55.8556C58.463 53.2481 60.5314 50.1526 61.9425 46.7458C62.8472 44.5617 63.4704 42.2771 63.8012 39.9469C64.1086 37.7816 65.853 36 68.04 36V36Z\" fill=\"var(--br-loading-color-foreground)\"/>"
    }
  };
  window.BRLoadingInit = function (raiz) {
    $all('.br-loading', raiz || d).forEach(function (el) {
      if (el._svgPronto || el.hasAttribute('data-progress')) return;
      el._svgPronto = true;
      var tam = el.classList.contains('large') ? 'large'
        : el.classList.contains('medium') ? 'medium' : 'small';
      // Sempre a arte "main": o traçado é o mesmo nos dois modos, e a cor
      // vem das custom properties, que o CSS troca no alto contraste.
      var art = LOADING_SVG[tam + '-main'];
      if (!art) return;
      el.innerHTML = '<svg class="loading-art" viewBox="' + art.vb +
        '" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' + art.d + '</svg>';
    });
  };

  window.BRMessageInit = function (raiz) {
    $all('.br-message', raiz || d).forEach(function (m) {
      if (m._tipoPronto) return;
      m._tipoPronto = true;
      var tipo = Object.keys(MSG_ROTULO).filter(function (t) { return m.classList.contains(t); })[0];
      if (!tipo) return;
      var alvo = m.querySelector('.icon') || m;
      var rot = d.createElement('span');
      rot.className = 'sr-only';
      rot.textContent = MSG_ROTULO[tipo] + ': ';
      alvo.insertBefore(rot, alvo.firstChild);
    });
  };
  window.BRTooltipInit = function (raiz) {
    $all('.br-tooltip', raiz || d).forEach(function (tt) {
      if (tt._pronto) return;
      tt._pronto = true;
      if (!tt.getAttribute('place')) tt.setAttribute('place', 'top');
      if (!tt.querySelector('.arrow')) {
        var seta = d.createElement('div');
        seta.className = 'arrow';
        seta.setAttribute('aria-hidden', 'true');
        tt.appendChild(seta);
      }
      var popover = tt.hasAttribute('popover');
      if (popover && !tt.querySelector('.close')) {
        var fechar = d.createElement('button');
        fechar.type = 'button';
        fechar.className = 'close';
        fechar.setAttribute('aria-label', 'Fechar');
        fechar.innerHTML = '<i class="fas fa-times" aria-hidden="true"></i>';
        fechar.addEventListener('click', function () { ttEsconder(tt); });
        tt.appendChild(fechar);
      }
      var gat = ttGatilho(tt);
      if (gat) {
        // O conteúdo do tooltip também descreve o gatilho, para o leitor de
        // tela — o balão em si nunca recebe foco.
        if (!tt.id) tt.id = 'tt-' + Math.random().toString(36).slice(2, 8);
        /* aria-describedby é o atributo que a diretriz pede: o balão
           COMPLEMENTA o gatilho, não o nomeia. Quando o gatilho não tem
           nome acessível próprio, aí sim o balão vira o rótulo. */
        if (!gat.getAttribute('aria-describedby')) gat.setAttribute('aria-describedby', tt.id);
        if (!gat.getAttribute('aria-labelledby') && !gat.getAttribute('aria-label') &&
            !(gat.textContent || '').trim()) {
          gat.setAttribute('aria-labelledby', tt.id);
        }
        if (popover) {
          gat.addEventListener('click', function () {
            if (tt.hasAttribute('data-show')) ttEsconder(tt); else ttMostrar(tt);
          });
        } else {
          ['mouseenter', 'focus'].forEach(function (e) {
            gat.addEventListener(e, function () { ttMostrar(tt); });
          });
          // O blur (saída por teclado) esconde na hora; o mouseleave passa
          // pela carência, para o balão continuar alcançável pelo cursor.
          gat.addEventListener('mouseleave', function () { ttAgendarSaida(tt); });
          // Gatilho que também abre menu (densidade da tabela, ações):
          // o balão sai de cena assim que o menu aparece, senão fica por
          // cima da primeira opção.
          gat.addEventListener('click', function () { ttEsconder(tt); });
          gat.addEventListener('blur', function () { ttEsconder(tt); });
          tt.addEventListener('mouseenter', function () { ttCancelarSaida(tt); });
          tt.addEventListener('mouseleave', function () { ttAgendarSaida(tt); });
        }
      }
      if (tt.hasAttribute('active')) ttMostrar(tt);
    });
  };
  d.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape') $all('.br-tooltip[data-show]').forEach(ttEsconder);
  });
  d.addEventListener('click', function (ev) {
    $all('.br-tooltip[popover][data-show]').forEach(function (tt) {
      var gat = ttGatilho(tt);
      if (!tt.contains(ev.target) && (!gat || !gat.contains(ev.target))) ttEsconder(tt);
    });
  });

  /* ── UPLOAD (br-upload) ── Monta o botão da área de transferência a
     partir do input[type=file], trata clique e arrastar-e-soltar (estado
     dropzone), exibe Loading por arquivo enquanto transfere, lista os
     arquivos com nome truncado + tooltip, tamanho formatado e botão
     excluir, e dá feedback quando o modo é de arquivo único. ── */
  function upTamanho(n) {
    /* Padrão Internacionalização: número formatado pela convenção local
       (Intl.NumberFormat) em vez de toFixed — em pt-BR o separador decimal
       é a vírgula, e toFixed devolvia sempre o ponto. */
    var fmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (n < 1024) return new Intl.NumberFormat('pt-BR').format(n) + ' bytes';
    var u = ['KB', 'MB', 'GB', 'TB'], i = 0, v = n / 1024, s = '';
    while (v >= 1 && i < u.length) { s = fmt.format(v) + ' ' + u[i]; v /= 1024; i++; }
    return s;
  }
  window.BRUploadInit = function (raiz) {
    $all('.br-upload', raiz || d).forEach(function (comp) {
      if (comp._pronto) return;
      comp._pronto = true;
      var input = comp.querySelector('.upload-input');
      var lista = comp.querySelector('.upload-list');
      if (!input || !lista) return;
      var multi = input.hasAttribute('multiple');
      var arquivos = [];
      var botao = d.createElement('button');
      botao.type = 'button';
      botao.className = 'upload-button';
      botao.innerHTML = '<i class="fas fa-upload" aria-hidden="true"></i><span>' +
        (multi ? 'Selecione o(s) arquivo(s)' : 'Selecione o arquivo') + '</span>';
      if (comp.hasAttribute('disabled')) botao.disabled = true;
      comp.insertBefore(botao, lista);

      function limparFeedback() {
        $all('.feedback', comp).forEach(function (f) { f.remove(); });
        ['success', 'danger', 'warning', 'info'].forEach(function (s) { comp.removeAttribute('data-' + s); });
      }
      function feedback(tipo, texto) {
        limparFeedback();
        var icones = { danger: 'fa-times-circle', warning: 'fa-exclamation-triangle',
          info: 'fa-info-circle', success: 'fa-check-circle' };
        comp.setAttribute('data-' + tipo, 'data-' + tipo);
        var m = d.createElement('span');
        m.className = 'feedback ' + tipo + ' mt-1';
        m.setAttribute('role', 'alert');
        m.setAttribute('aria-live', 'assertive');
        m.innerHTML = '<i class="fas ' + icones[tipo] + '" aria-hidden="true"></i>' + texto;
        lista.parentNode.insertBefore(m, lista);
      }
      function desenhar() {
        lista.innerHTML = arquivos.map(function (f, i) {
          return '<div class="br-item"><span class="content tooltip-wrap" data-tooltip-trigger tabindex="0">' +
            f.name + '<span class="br-tooltip small" role="tooltip" info place="top">' +
            '<span class="subtext">' + f.name + '</span></span></span>' +
            '<span class="support"><span>' + upTamanho(f.size) + '</span>' +
            '<button class="br-button circle small" type="button" data-remover="' + i +
            '" aria-label="Apagar arquivo ' + f.name + '"><i class="fas fa-trash-alt" aria-hidden="true"></i></button>' +
            '</span></div>';
        }).join('');
        window.BRTooltipInit(lista);
      }
      function receber(files) {
        limparFeedback();
        var novos = Array.prototype.slice.call(files);
        if (!novos.length) return;
        if (!multi && novos.length > 1) return feedback('danger', 'É permitido o envio de somente 1 arquivo.');
        var substituiu = !multi && arquivos.length > 0;
        if (!multi) arquivos = [];
        arquivos = arquivos.concat(novos);
        // Loading (dependência da anatomia) enquanto a transferência ocorre.
        lista.innerHTML = '<div class="br-loading medium" role="progressbar" aria-label="Enviando arquivos"></div>';
        // Criado depois do init da tela: recebe a arte oficial aqui.
        if (window.BRLoadingInit) window.BRLoadingInit(lista);
        setTimeout(function () {
          desenhar();
          if (substituiu) feedback('warning', 'O arquivo enviado anteriormente foi substituído.');
          else feedback('success', novos.length === 1 ? 'Arquivo anexado.' : novos.length + ' arquivos anexados.');
        }, 500);
      }
      botao.addEventListener('click', function () { input.click(); });
      input.addEventListener('change', function () { receber(input.files); });
      lista.addEventListener('click', function (ev) {
        var b = ev.target.closest('[data-remover]');
        if (!b) return;
        arquivos.splice(+b.getAttribute('data-remover'), 1);
        limparFeedback();
        desenhar();
        input.value = '';
      });
      ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (e) {
        botao.addEventListener(e, function (ev) { ev.preventDefault(); ev.stopPropagation(); });
      });
      ['dragenter', 'dragover'].forEach(function (e) {
        botao.addEventListener(e, function () { comp.classList.add('dragging'); });
      });
      ['dragleave', 'drop'].forEach(function (e) {
        botao.addEventListener(e, function () { comp.classList.remove('dragging'); });
      });
      botao.addEventListener('drop', function (ev) { receber(ev.dataTransfer.files); });
      if (comp.hasAttribute('disabled')) feedback('warning', 'Upload desabilitado');
    });
  };

  /* ── WIZARD (br-wizard) ── Numera as etapas, aplica a etapa inicial e
     liga o painel de etapas à área de conteúdo. A barra de navegação segue
     a diretriz: Voltar some na primeira etapa e Avançar vira Concluir na
     última. ── */
  window.BRWizardInit = function (raiz) {
    $all('.br-wizard', raiz || d).forEach(function (comp) {
      if (comp._pronto) return;
      comp._pronto = true;
      var botoes = $all('.wizard-progress-btn', comp);
      var paineis = $all('.wizard-panel', comp);
      botoes.forEach(function (b, i) { b.setAttribute('step', i + 1); });
      function ir(n) {
        n = Math.max(0, Math.min(paineis.length - 1, n));
        botoes.forEach(function (b, i) {
          var on = i === n;
          if (on) b.setAttribute('active', ''); else b.removeAttribute('active');
          b.setAttribute('aria-selected', String(on));
        });
        paineis.forEach(function (p, i) {
          if (i === n) p.setAttribute('active', ''); else p.removeAttribute('active');
        });
        var atual = paineis[n];
        if (atual) {
          var prev = atual.querySelector('.wizard-btn-prev');
          if (prev) prev.hidden = n === 0;
          var next = atual.querySelector('.wizard-btn-next');
          if (next) next.textContent = n === paineis.length - 1 ? 'Concluir' : 'Avançar';
          var alvo = atual.querySelector('.wizard-panel-content');
          if (alvo) alvo.focus({ preventScroll: true });
        }
      }
      comp.addEventListener('click', function (ev) {
        var b = ev.target.closest('.wizard-progress-btn');
        if (b && !b.disabled) return ir(botoes.indexOf(b));
        var nav = ev.target.closest('.wizard-btn-next, .wizard-btn-prev');
        if (!nav) return;
        var painel = nav.closest('.wizard-panel');
        var i = paineis.indexOf(painel);
        ir(nav.classList.contains('wizard-btn-prev') ? i - 1 : i + 1);
      });
      ir(Math.max(0, (+comp.getAttribute('step') || 1) - 1));
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
    iniciarCarousels: iniciarCarousels,
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
