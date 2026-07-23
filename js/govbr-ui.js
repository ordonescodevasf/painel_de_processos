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
  d.addEventListener('click', function (ev) {
    var btn = ev.target.closest('[data-toggle="dropdown"]');
    if (btn) {
      var panel = d.getElementById(btn.getAttribute('data-target'));
      if (!panel) return;
      var abrir = panel.hidden;
      closeAllDropdowns(abrir ? panel : null);
      panel.hidden = !abrir;
      btn.setAttribute('aria-expanded', abrir ? 'true' : 'false');
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
