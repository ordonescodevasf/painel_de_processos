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

  /* ── Cookies: preferência mínima (este painel não usa cookies de
        rastreio; só localStorage para contraste/preferências) ──────── */
  d.addEventListener('click', function (ev) {
    if (ev.target.closest('#btnCookieSettings')) {
      var m = $('#modalCookies'); if (m) m.hidden = false;
    }
  });

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
    var downloadEl = d.getElementById('reportErrorShotDownload');
    var textEl = d.getElementById('reportErrorText');
    var shotDataUrl = null;

    function setStatus(msg) { statusEl.textContent = msg; }

    function captureScreenshot() {
      shotDataUrl = null;
      previewEl.setAttribute('hidden', '');
      downloadEl.setAttribute('hidden', '');
      if (!navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) {
        setStatus('Este navegador não permite captura automática de tela. Tire um print manualmente (Windows: tecla Print Screen ou Win+Shift+S; Mac: Cmd+Shift+4) e anexe ao e-mail.');
        return;
      }
      setStatus('Capturando print da tela… escolha a tela/aba na janela que o navegador abrir.');
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
          previewEl.removeAttribute('hidden');
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

  /* ── COOKIEBAR (simplificado) — persistência via localStorage; só é
     dispensado por ação explícita ("Entendi"), sem Esc/clique fora. ── */
  (function () {
    var KEY = 'painel_processos_cookie_notice_ack_v1';
    var notice = d.getElementById('cookieNotice');
    var scrim = d.getElementById('cookieScrim');
    if (!notice) return;
    var acceptBtn = d.getElementById('cookieNoticeAccept');
    var detailsBtn = d.getElementById('cookieNoticeDetails');
    var detailsBox = d.getElementById('cookieNoticeDetailBox');

    function jaAceito() {
      try { return !!localStorage.getItem(KEY); } catch (e) { return false; }
    }
    function release() {
      d.body.style.overflow = '';
      notice.classList.add('d-none');
      if (scrim) scrim.setAttribute('hidden', '');
      try { localStorage.setItem(KEY, '1'); } catch (e) { /* não persiste entre sessões */ }
    }
    if (acceptBtn) acceptBtn.addEventListener('click', release);
    if (detailsBtn && detailsBox) {
      detailsBtn.addEventListener('click', function () {
        var open = detailsBtn.getAttribute('aria-expanded') === 'true';
        detailsBtn.setAttribute('aria-expanded', String(!open));
        if (open) detailsBox.setAttribute('hidden', ''); else detailsBox.removeAttribute('hidden');
      });
    }
    if (jaAceito()) { notice.classList.add('d-none'); if (scrim) scrim.setAttribute('hidden', ''); }
  })();

  /* ── ONBOARDING — tour de apresentação em carrossel; abre sozinho no
     primeiro acesso (a menos que já dispensado) e pode ser reaberto por
     "Como usar este painel" (rodapé/menu) ou pelo switch dedicado. ── */
  (function () {
    var scrimEl = d.getElementById('onboardingScrim');
    if (!scrimEl) return;
    var col = d.getElementById('onboardingCol');
    var pages = Array.prototype.slice.call(col.children);
    var total = pages.length;
    var nav = scrimEl.querySelector('.ob-nav');
    var stepText = d.getElementById('onboardingStepText');
    var prevBtn = scrimEl.querySelector('.ob-prev');
    var nextBtn = scrimEl.querySelector('.ob-next');
    var startBtn = scrimEl.querySelector('.ob-start');
    var finishBtn = scrimEl.querySelector('.ob-finish');
    var skipLink = d.getElementById('onboardingSkip');
    var autoToggle = d.getElementById('obAutoToggle');
    var current = 0;

    function render() {
      pages.forEach(function (p, i) { p.classList.toggle('active', i === current); });
      var isFirst = current === 0, isLast = current === total - 1;
      nav.style.display = (isFirst || isLast) ? 'none' : 'flex';
      stepText.textContent = (isFirst || isLast) ? '' : ('Passo ' + current + ' de ' + (total - 2));
    }
    function go(idx) { current = Math.max(0, Math.min(total - 1, idx)); render(); }
    prevBtn.addEventListener('click', function () { go(current - 1); });
    nextBtn.addEventListener('click', function () { go(current + 1); });
    if (startBtn) startBtn.addEventListener('click', function () { go(1); });
    function fechar() { scrimEl.hidden = true; }
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
      autoToggle.addEventListener('change', function () { setDismissed(!autoToggle.checked); });
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

  /* ── API mínima para o app preencher o menu de seções ────────────── */
  window.PPUI = {
    setMenuSections: function (itens) {
      var ul = $('#sectionMenuList');
      if (!ul) return;
      ul.innerHTML = itens.map(function (it) {
        return '<li><a class="menu-item" href="' + it.href + '">' +
          '<span class="icon"><i class="fas ' + it.icone + '" aria-hidden="true"></i></span>' +
          '<span>' + it.rotulo + '</span></a></li>';
      }).join('');
      ul.addEventListener('click', function () {
        var m = $('#sectionMenu'); if (m) m.classList.remove('active');
      });
    },
    fecharMenus: function () { closeAllDropdowns(null); }
  };
})();
