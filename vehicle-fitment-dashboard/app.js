/* Vehicle fitment dashboard — rendering logic.
   Data comes from data.js as window.VLDATA. No build step, no dependencies. */
(function () {
  'use strict';

  var D = window.VLDATA;
  var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TYPES = [
    { k: 'n', label: 'Non EV',            col: 'var(--t-non)', hex: '#1F3D4A', cls: 'n1b' },
    { k: 'e', label: 'EV',                col: 'var(--t-ev)',  hex: '#1B84D8', cls: 'e1'  },
    { k: 'u', label: 'Unknown powertrain', col: 'var(--t-unk)', hex: '#C0399F', cls: 'u1' }
  ];

  // ---------- helpers ----------
  function nice(v) {
    var p = Math.pow(10, Math.floor(Math.log(v) / Math.LN10));
    var steps = [1, 1.5, 2, 2.5, 3, 4, 5, 7.5, 10];
    for (var i = 0; i < steps.length; i++) { if (p * steps[i] >= v) return p * steps[i]; }
    return p * 10;
  }
  function fmt(n) { return n.toLocaleString('en-AU'); }
  function shortNum(n) {
    if (n >= 1000) return (n % 1000 === 0 ? n / 1000 : (n / 1000).toFixed(1)) + 'k';
    return String(n);
  }
  function median(arr) {
    var s = arr.slice().sort(function (a, b) { return a - b; });
    return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
  }
  function monthLabel(m) { var p = m.split('-'); return MN[parseInt(p[1], 10) - 1] + ' ' + p[0]; }
  function el(tag, cls) { var e = document.createElement(tag); if (cls) e.className = cls; return e; }
  function svgEl(tag) { return document.createElementNS('http://www.w3.org/2000/svg', tag); }

  /* Draw gridlines and y labels into a plot frame.
     isPct switches the label suffix. Returns nothing. */
  function drawAxis(gridNode, labNode, cap, isPct) {
    var step = nice(cap / 5);
    gridNode.innerHTML = '';
    labNode.innerHTML = '';
    for (var v = 0; v <= cap + 1e-9; v += step) {
      var pos = 100 - v / cap * 100;
      var g = el('i'); g.style.top = pos + '%';
      if (v === 0) g.className = 'zero';
      gridNode.appendChild(g);
      var s = el('span'); s.style.top = pos + '%';
      s.textContent = isPct ? (Math.round(v * 10) / 10) + '%' : shortNum(Math.round(v));
      labNode.appendChild(s);
    }
  }

  function drawTicks(node, view, key) {
    node.innerHTML = '';
    var every = view.length > 36 ? 6 : (view.length > 24 ? 3 : (view.length > 14 ? 2 : 1));
    view.forEach(function (d, i) {
      var span = el('span');
      if (i % every === 0 || view.length <= 14) {
        var p = d[key].split('-');
        span.innerHTML = MN[parseInt(p[1], 10) - 1] + '<b>' + p[0] + '</b>';
      }
      node.appendChild(span);
    });
  }

  // ---------- chart 1: vehicles logged per month ----------
  var monthly = (function () {
    var chart = document.getElementById('chart'),
        ticks = document.getElementById('ticks'),
        grid  = document.getElementById('grid'),
        ylab  = document.getElementById('ylab'),
        note  = document.getElementById('chartnote'),
        donut = document.getElementById('donut'),
        body  = document.getElementById('drillbody'),
        hint  = document.getElementById('drillhint'),
        sub   = document.getElementById('drillsub'),
        dtit  = document.getElementById('drilltitle'),
        dleg  = document.getElementById('drilllegend');

    var range = 12, mode = 'clip', selected = null;
    var LAST = D.months[D.months.length - 1].m;

    function drawDonut(row) {
      var total = row.e + row.n + row.u, R = 52, CX = 70, CY = 70,
          C = 2 * Math.PI * R, off = 0;
      var out = '<circle cx="' + CX + '" cy="' + CY + '" r="' + R +
                '" fill="none" stroke="#E4E9E7" stroke-width="22"/>';
      TYPES.forEach(function (t) {
        var v = row[t.k];
        if (!v) return;
        var len = v / total * C;
        out += '<circle cx="' + CX + '" cy="' + CY + '" r="' + R + '" fill="none" stroke="' + t.hex +
               '" stroke-width="22" stroke-dasharray="' + len.toFixed(2) + ' ' + (C - len).toFixed(2) +
               '" stroke-dashoffset="' + (-off).toFixed(2) + '" transform="rotate(-90 ' + CX + ' ' + CY +
               ')"><title>' + t.label + ': ' + fmt(v) + '</title></circle>';
        off += len;
      });
      out += '<text x="70" y="66" text-anchor="middle" font-size="19" font-weight="700" fill="#16262B">' +
             shortNum(total) + '</text>' +
             '<text x="70" y="82" text-anchor="middle" font-size="9" fill="#4A5F63">unfitted</text>';
      donut.innerHTML = '<title id="donuttitle">Unfitted vehicles by type</title>' + out;
    }

    function select(m) {
      selected = (m === null || selected === m) ? null : m;
      if (!selected) {
        body.hidden = true; hint.hidden = false; sub.hidden = true;
      } else {
        var row = D.months.filter(function (d) { return d.m === selected; })[0];
        var total = row.e + row.n + row.u;
        drawDonut(row);
        dtit.textContent = monthLabel(selected) + ' \u2014 ' + fmt(total) + ' of ' +
                           fmt(row.w + row.o) + ' vehicles logged had no tyre sizes';
        dleg.innerHTML = TYPES.map(function (t) {
          var v = row[t.k];
          return '<div class="dleg"><span class="sw2" style="background:' + t.hex + '"></span>' +
                 '<span>' + t.label + '</span><span class="n1">' + fmt(v) + '</span>' +
                 '<span class="n2">' + (total ? Math.round(v / total * 100) : 0) + '%</span></div>';
        }).join('');
        body.hidden = false; hint.hidden = true; sub.hidden = false;
      }
      render();
    }

    function render() {
      var view = range > 0 ? D.months.slice(Math.max(0, D.months.length - range)) : D.months.slice();
      var totals = view.map(function (d) { return d.w + d.o; });
      var peak = Math.max.apply(null, totals);
      var cap = mode === 'full' ? nice(peak) : nice(Math.max(median(totals) * 2.5, 1000));
      if (cap > peak) cap = nice(peak);
      var over = view.filter(function (d) { return d.w + d.o > cap; });

      drawAxis(grid, ylab, cap, false);
      chart.innerHTML = '';

      view.forEach(function (d) {
        var tot = d.w + d.o, clipped = tot > cap, k = clipped ? cap / tot : 1;
        var col = el('div', 'col pick' + (clipped ? ' clip' : '') +
                            (d.m === LAST ? ' faint' : '') + (d.m === selected ? ' sel' : ''));
        col.title = monthLabel(d.m) + '  \u2014  ' + fmt(tot) + ' logged\n' +
                    fmt(d.w) + ' with sizes (' + Math.round(d.w / tot * 100) + '%)\n' +
                    fmt(d.o) + ' without' + (d.m === LAST ? '\n(partial month)' : '');
        col.setAttribute('role', 'button');
        col.setAttribute('tabindex', '0');
        col.setAttribute('aria-label', monthLabel(d.m) + ', show unfitted vehicles by type');
        if (clipped) { var u = el('u'); u.textContent = fmt(tot); col.appendChild(u); }
        var o = el('i', 'wo'); o.style.height = (d.o / cap * 100 * k).toFixed(2) + '%';
        var w = el('i', 'w');  w.style.height = (d.w / cap * 100 * k).toFixed(2) + '%';
        col.appendChild(o); col.appendChild(w);
        col.addEventListener('click', function () { select(d.m); });
        col.addEventListener('keydown', function (ev) {
          if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); select(d.m); }
        });
        chart.appendChild(col);
      });

      drawTicks(ticks, view, 'm');

      var msg = 'Axis tops out at ' + fmt(cap) + ' vehicles.';
      if (over.length) {
        msg += ' ' + over.map(function (d) { return monthLabel(d.m); }).join(' and ') +
               (over.length > 1 ? ' run' : ' runs') +
               ' past it, drawn hatched with the true total above. Switch to full scale to see them whole.';
      }
      msg += ' ' + monthLabel(LAST) + ' is a partial month.';
      note.textContent = msg;
    }

    document.getElementById('drillclear').addEventListener('click', function () { select(null); });
    return { render: render, setRange: function (n) { range = n; }, setMode: function (m) { mode = m; } };
  })();

  // ---------- chart 2: composition of the unfitted pile ----------
  var trend = (function () {
    var chart = document.getElementById('tchart'),
        lines = document.getElementById('tlines'),
        grid  = document.getElementById('tgrid'),
        ylab  = document.getElementById('tylab'),
        ticks = document.getElementById('tticks'),
        leg   = document.getElementById('tlegend'),
        axis  = document.getElementById('taxis');

    var mode = 'share', range = 0;

    function renderShare(view) {
      var pts = { e: [], n: [], u: [] };
      view.forEach(function (d) {
        var t = d.e + d.n + d.u || 1;
        pts.e.push(d.e / t * 100); pts.n.push(d.n / t * 100); pts.u.push(d.u / t * 100);
      });
      var cap = 100;
      drawAxis(grid, ylab, cap, true);

      var W = 1000, H = 250, n = view.length;
      function px(i) { return n === 1 ? W / 2 : (i / (n - 1)) * W; }
      function py(v) { return H - v / cap * H; }

      [['n', '#1F3D4A'], ['u', '#C0399F'], ['e', '#1B84D8']].forEach(function (pair) {
        var k = pair[0], col = pair[1], d = '';
        pts[k].forEach(function (v, i) { d += (i ? 'L' : 'M') + px(i).toFixed(1) + ' ' + py(v).toFixed(1) + ' '; });
        var path = svgEl('path');
        path.setAttribute('d', d);
        path.setAttribute('stroke', col);
        lines.appendChild(path);

        var name = k === 'e' ? 'EV' : (k === 'n' ? 'Non EV' : 'Unknown powertrain');
        pts[k].forEach(function (v, i) {
          if (n > 40 && i % 2) return;
          var c = svgEl('circle');
          c.setAttribute('cx', px(i)); c.setAttribute('cy', py(v));
          c.setAttribute('r', n > 30 ? 2.5 : 3.5);
          c.setAttribute('fill', col);
          var ti = svgEl('title');
          ti.textContent = monthLabel(view[i].m) + '  ' + name + ' ' + (Math.round(v * 10) / 10) +
                           '% of ' + fmt(view[i].e + view[i].n + view[i].u) + ' unfitted';
          c.appendChild(ti);
          lines.appendChild(c);
        });
      });

      leg.innerHTML = TYPES.map(function (t) {
        return '<div><span class="sw" style="background:' + t.hex + '"></span><em>' + t.label + '</em></div>';
      }).join('') +
      '<div><em>Read as: of the vehicles missing tyre sizes that month, this share was that type ' +
      '\u2014 not the share of that type which is unfitted. The three lines sum to 100%.</em></div>';
      axis.textContent = 'Vertical axis: each type as a percentage of that month\u2019s unfitted vehicles';
    }

    function renderCount(view) {
      var totals = view.map(function (d) { return d.e + d.n + d.u; });
      var peak = Math.max.apply(null, totals);
      var cap = nice(Math.max(median(totals) * 2.5, 1000));
      if (cap > peak) cap = nice(peak);
      drawAxis(grid, ylab, cap, false);

      view.forEach(function (d) {
        var tot = d.e + d.n + d.u, clipped = tot > cap, k = clipped ? cap / tot : 1;
        var col = el('div', 'col' + (clipped ? ' clip' : ''));
        col.title = monthLabel(d.m) + '  \u2014  ' + fmt(tot) + ' unfitted\n' +
                    'Non EV ' + fmt(d.n) + '\nEV ' + fmt(d.e) + '\nUnknown ' + fmt(d.u);
        TYPES.slice().reverse().forEach(function (t) {
          var seg = el('i', t.cls);
          seg.style.height = (d[t.k] / cap * 100 * k).toFixed(2) + '%';
          col.appendChild(seg);
        });
        if (clipped) { var u = el('u'); u.textContent = fmt(tot); col.appendChild(u); }
        chart.appendChild(col);
      });

      leg.innerHTML = TYPES.map(function (t) {
        return '<div><span class="sw" style="background:' + t.hex + '"></span><em>' + t.label + '</em></div>';
      }).join('') +
      '<div><em>Axis tops out at ' + fmt(cap) + '; taller months are hatched with their true total.</em></div>';
      axis.textContent = 'Vertical axis: number of unfitted vehicles that month, stacked by type';
    }

    function render() {
      var view = range > 0 ? D.months.slice(Math.max(0, D.months.length - range)) : D.months.slice();
      chart.innerHTML = '';
      lines.innerHTML = '';
      if (mode === 'share') renderShare(view); else renderCount(view);
      drawTicks(ticks, view, 'm');
    }
    return { render: render, setMode: function (m) { mode = m; }, setRange: function (n) { range = n; } };
  })();

  // ---------- control wiring ----------
  function wire(selector, apply, redraw) {
    var btns = document.querySelectorAll(selector);
    Array.prototype.forEach.call(btns, function (b) {
      b.addEventListener('click', function () {
        Array.prototype.forEach.call(btns, function (x) { x.classList.remove('on'); });
        b.classList.add('on');
        apply(b);
        redraw();
      });
    });
  }
  wire('#monthly .ctrl button[data-n]', function (b) { monthly.setRange(parseInt(b.dataset.n, 10)); }, monthly.render);
  wire('#monthly .ctrl button[data-s]', function (b) { monthly.setMode(b.dataset.s); }, monthly.render);
  wire('#trend .ctrl button[data-t]',   function (b) { trend.setMode(b.dataset.t); },   trend.render);
  wire('#trend .ctrl button[data-r]',   function (b) { trend.setRange(parseInt(b.dataset.r, 10)); }, trend.render);

  monthly.render();
  trend.render();
})();
