/* ──────────────────────────────────────────────────────────────────
   DATETIMEPICKER (br-datetimepicker) — Componente gov.br DS 3.7.0.

   Implementação própria, sem a dependência do flatpickr usada pelo
   core oficial (o painel roda sem empacotador e sem node_modules).
   Cobre a anatomia completa do componente: campo de entrada, mês/ano
   com navegação (setas, select de mês e input de ano), dias da semana
   abreviados, dia atual em destaque, dia selecionado, hover, card,
   seletor de horas (24h), dias do mês anterior/posterior, datas
   inicial e final, intervalo em destaque e dias desabilitados; e os
   três tipos previstos — datepicker (data-type="text"), timepicker
   (data-type="time") e datetimepicker (data-type="datetime-local"),
   nos modos single e range (data-mode).

   Acessibilidade: abre com Enter, fecha com Esc, navega com Tab e
   setas, e cada dia é verbalizado por extenso ("sábado, 5 de agosto
   de 2023") conforme a recomendação de código do padrão.
   ────────────────────────────────────────────────────────────────── */
(function () {
  'use strict';
  var MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  var DIAS_ABREV = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
  var DIAS_EXTENSO = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira',
    'quinta-feira', 'sexta-feira', 'sábado'];

  function pad(n) { return n < 10 ? '0' + n : String(n); }
  function fmt(dt) { return pad(dt.getDate()) + '/' + pad(dt.getMonth() + 1) + '/' + dt.getFullYear(); }
  function parse(txt) {
    var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(txt || '').trim());
    if (!m) return null;
    var dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(dt) ? null : dt;
  }
  function mesmoDia(a, b) {
    return !!a && !!b && a.getFullYear() === b.getFullYear() &&
      a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }
  function porExtenso(dt) {
    return DIAS_EXTENSO[dt.getDay()] + ', ' + dt.getDate() + ' de ' +
      MESES[dt.getMonth()].toLowerCase() + ' de ' + dt.getFullYear();
  }

  function Picker(root) {
    this.root = root;
    this.input = root.querySelector('input');
    this.botao = root.querySelector('.br-input .br-button');
    this.modo = root.getAttribute('data-mode') || 'single';
    this.tipo = root.getAttribute('data-type') || 'text';
    this.comCalendario = this.tipo !== 'time';
    this.comHora = this.tipo === 'time' || this.tipo === 'datetime-local';
    this.min = parse(root.getAttribute('data-min-date'));
    this.max = parse(root.getAttribute('data-max-date'));
    this.inicio = null;
    this.fim = null;
    this.hora = 0;
    this.minuto = 0;
    this.cursor = new Date();
    this._montar();
    this._ler();
    this._eventos();
  }

  Picker.prototype._montar = function () {
    var c = document.createElement('div');
    c.className = 'dtp-calendar';
    c.hidden = true;
    c.setAttribute('role', 'dialog');
    c.setAttribute('aria-label', this.comCalendario ? 'Selecionar data' : 'Selecionar hora');
    c.innerHTML =
      (this.comCalendario
        ? '<div class="dtp-months">' +
        '<button class="br-button circle small terciary dtp-prev" type="button" aria-label="Mês anterior"><i class="fas fa-angle-left" aria-hidden="true"></i></button>' +
        '<select class="dtp-month-select" aria-label="Mês">' +
        MESES.map(function (m, i) { return '<option value="' + i + '">' + m + '</option>'; }).join('') +
        '</select>' +
        '<input class="dtp-year" type="number" inputmode="numeric" min="1900" max="2999" aria-label="Ano">' +
        '<button class="br-button circle small terciary dtp-next" type="button" aria-label="Próximo mês"><i class="fas fa-angle-right" aria-hidden="true"></i></button>' +
        '</div>' +
        '<div class="dtp-weekdays" aria-hidden="true">' +
        DIAS_ABREV.map(function (d) { return '<span>' + d + '</span>'; }).join('') + '</div>' +
        '<div class="dtp-days" role="grid"></div>'
        : '') +
      (this.comHora
        ? '<div class="dtp-time">' +
        '<div class="dtp-num"><button class="br-button circle small terciary" type="button" data-step="h1" aria-label="Aumentar hora"><i class="fas fa-chevron-up" aria-hidden="true"></i></button>' +
        '<input class="dtp-hora" type="number" min="0" max="23" aria-label="Hora">' +
        '<button class="br-button circle small terciary" type="button" data-step="h-1" aria-label="Diminuir hora"><i class="fas fa-chevron-down" aria-hidden="true"></i></button></div>' +
        '<span class="dtp-sep">:</span>' +
        '<div class="dtp-num"><button class="br-button circle small terciary" type="button" data-step="m1" aria-label="Aumentar minuto"><i class="fas fa-chevron-up" aria-hidden="true"></i></button>' +
        '<input class="dtp-minuto" type="number" min="0" max="59" aria-label="Minuto">' +
        '<button class="br-button circle small terciary" type="button" data-step="m-1" aria-label="Diminuir minuto"><i class="fas fa-chevron-down" aria-hidden="true"></i></button></div>' +
        '</div>'
        : '');
    this.root.appendChild(c);
    this.cal = c;
    this.grade = c.querySelector('.dtp-days');
    this.selMes = c.querySelector('.dtp-month-select');
    this.inpAno = c.querySelector('.dtp-year');
    this.inpHora = c.querySelector('.dtp-hora');
    this.inpMinuto = c.querySelector('.dtp-minuto');
  };

  // Lê o valor já presente no input (formato dd/mm/aaaa [hh:mm])
  Picker.prototype._ler = function () {
    var v = String(this.input.value || '');
    if (this.tipo === 'time') {
      var t = /^(\d{1,2}):(\d{2})$/.exec(v.trim());
      if (t) { this.hora = Number(t[1]); this.minuto = Number(t[2]); }
      return;
    }
    var partes = v.split(' até ');
    this.inicio = parse(partes[0]);
    if (this.modo === 'range') this.fim = parse(partes[1]);
    var h = /(\d{2}):(\d{2})/.exec(v);
    if (h) { this.hora = Number(h[1]); this.minuto = Number(h[2]); }
    if (this.inicio) this.cursor = new Date(this.inicio.getTime());
  };

  Picker.prototype._desabilitado = function (dt) {
    if (this.min && dt < new Date(this.min.getFullYear(), this.min.getMonth(), this.min.getDate())) return true;
    if (this.max && dt > new Date(this.max.getFullYear(), this.max.getMonth(), this.max.getDate())) return true;
    // Ao escolher a data final, tudo antes da inicial fica indisponível
    if (this.modo === 'range' && this.inicio && !this.fim && dt < this.inicio) return true;
    return false;
  };

  Picker.prototype.render = function () {
    if (this.comHora) {
      this.inpHora.value = pad(this.hora);
      this.inpMinuto.value = pad(this.minuto);
    }
    if (!this.comCalendario) return;
    this.selMes.value = String(this.cursor.getMonth());
    this.inpAno.value = String(this.cursor.getFullYear());
    var ano = this.cursor.getFullYear(), mes = this.cursor.getMonth();
    var primeiro = new Date(ano, mes, 1);
    var inicioGrade = new Date(ano, mes, 1 - primeiro.getDay());
    var hoje = new Date();
    var html = '';
    for (var i = 0; i < 42; i++) {
      var dt = new Date(inicioGrade.getFullYear(), inicioGrade.getMonth(), inicioGrade.getDate() + i);
      var cls = ['dtp-day'];
      if (dt.getMonth() !== mes) cls.push('other-month');
      if (mesmoDia(dt, hoje)) cls.push('today');
      var sel = mesmoDia(dt, this.inicio) || mesmoDia(dt, this.fim);
      if (sel) cls.push('selected');
      else if (this.inicio && this.fim && dt > this.inicio && dt < this.fim) cls.push('in-range');
      var off = this._desabilitado(dt);
      html += '<button type="button" class="' + cls.join(' ') + '" data-data="' + fmt(dt) + '"' +
        (off ? ' disabled' : '') + ' aria-label="' + porExtenso(dt) + '"' +
        (sel ? ' aria-current="date"' : '') + '>' + dt.getDate() + '</button>';
    }
    this.grade.innerHTML = html;
  };

  Picker.prototype._valor = function () {
    if (this.tipo === 'time') return pad(this.hora) + ':' + pad(this.minuto);
    if (!this.inicio) return '';
    var txt = fmt(this.inicio);
    if (this.modo === 'range') txt += this.fim ? ' até ' + fmt(this.fim) : '';
    if (this.tipo === 'datetime-local') txt += ' ' + pad(this.hora) + ':' + pad(this.minuto);
    return txt;
  };
  Picker.prototype._aplicar = function (completo) {
    this.input.value = this._valor();
    // No modo intervalo o change só dispara com as duas datas escolhidas:
    // quem escuta costuma redesenhar a tela, o que interromperia a
    // seleção da data final.
    if (completo !== false) this.input.dispatchEvent(new Event('change', { bubbles: true }));
  };

  Picker.prototype.abrir = function () {
    if (!this.cal.hidden) return;
    this.render();
    this.cal.hidden = false;
    this.input.setAttribute('aria-expanded', 'true');
  };
  Picker.prototype.fechar = function () {
    this.cal.hidden = true;
    this.input.setAttribute('aria-expanded', 'false');
  };

  Picker.prototype._eventos = function () {
    var self = this;
    this.input.setAttribute('aria-expanded', 'false');
    this.input.setAttribute('aria-haspopup', 'dialog');
    if (this.botao) {
      this.botao.removeAttribute('tabindex');
      this.botao.removeAttribute('aria-hidden');
      this.botao.addEventListener('click', function () {
        if (self.cal.hidden) self.abrir(); else self.fechar();
      });
    }
    this.input.addEventListener('click', function () { self.abrir(); });
    this.input.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') { ev.preventDefault(); self.abrir(); }
      if (ev.key === 'Escape') self.fechar();
    });
    // Máscara de digitação (dd/mm/aaaa, com separador de intervalo e hora)
    this.input.addEventListener('input', function () {
      if (self.tipo === 'time') return;
      var limpo = self.input.value;
      if (/[^\d/: até]/.test(limpo)) self.input.value = limpo.replace(/[^\d/: até]/g, '');
      self._ler();
    });
    this.input.addEventListener('blur', function () { self._ler(); });

    this.cal.addEventListener('click', function (ev) {
      var dia = ev.target.closest('.dtp-day');
      if (dia && !dia.disabled) {
        var dt = parse(dia.getAttribute('data-data'));
        if (self.modo === 'range') {
          if (!self.inicio || self.fim) { self.inicio = dt; self.fim = null; }
          else if (dt < self.inicio) { self.inicio = dt; }
          else { self.fim = dt; }
        } else self.inicio = dt;
        self.cursor = new Date(dt.getTime());
        self._aplicar(self.modo === 'single' ? true : !!self.fim);
        self.render();
        if (self.modo === 'single' && !self.comHora) self.fechar();
        if (self.modo === 'range' && self.fim) self.fechar();
        return;
      }
      var passo = ev.target.closest('[data-step]');
      if (passo) {
        var p = passo.getAttribute('data-step');
        if (p[0] === 'h') self.hora = (self.hora + Number(p.slice(1)) + 24) % 24;
        else self.minuto = (self.minuto + Number(p.slice(1)) + 60) % 60;
        self.render(); self._aplicar();
        return;
      }
      if (ev.target.closest('.dtp-prev')) { self.cursor.setMonth(self.cursor.getMonth() - 1); self.render(); }
      if (ev.target.closest('.dtp-next')) { self.cursor.setMonth(self.cursor.getMonth() + 1); self.render(); }
    });
    this.cal.addEventListener('change', function (ev) {
      if (ev.target === self.selMes) { self.cursor.setMonth(Number(self.selMes.value)); self.render(); }
      if (ev.target === self.inpAno) {
        var a = Number(self.inpAno.value);
        if (a >= 1900 && a <= 2999) { self.cursor.setFullYear(a); self.render(); }
      }
      if (ev.target === self.inpHora) { self.hora = Math.min(23, Math.max(0, Number(self.inpHora.value) || 0)); self.render(); self._aplicar(); }
      if (ev.target === self.inpMinuto) { self.minuto = Math.min(59, Math.max(0, Number(self.inpMinuto.value) || 0)); self.render(); self._aplicar(); }
    });
    // Navegação por teclado entre os dias
    this.cal.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape') { self.fechar(); self.input.focus(); return; }
      var dia = ev.target.closest('.dtp-day');
      if (!dia) return;
      var delta = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -7, ArrowDown: 7 }[ev.key];
      if (!delta) return;
      ev.preventDefault();
      var dt = parse(dia.getAttribute('data-data'));
      dt.setDate(dt.getDate() + delta);
      if (dt.getMonth() !== self.cursor.getMonth() || dt.getFullYear() !== self.cursor.getFullYear()) {
        self.cursor = new Date(dt.getTime());
        self.render();
      }
      var alvo = self.grade.querySelector('[data-data="' + fmt(dt) + '"]');
      if (alvo) alvo.focus();
    });
    document.addEventListener('click', function (ev) {
      if (!self.root.contains(ev.target)) self.fechar();
    });
  };

  window.PPDateTimePicker = {
    init: function (escopo) {
      (escopo || document).querySelectorAll('.br-datetimepicker').forEach(function (el) {
        if (el.__dtp) return;
        el.__dtp = new Picker(el);
      });
    }
  };
  document.addEventListener('DOMContentLoaded', function () { window.PPDateTimePicker.init(document); });
})();
