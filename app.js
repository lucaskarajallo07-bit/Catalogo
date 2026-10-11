/* La Plaza Perfumería: tienda estática.
   Datos: data/catalogo.json, generado desde el Excel con herramientas/excel_a_json.py. */
(function () {
  'use strict';

  var DATA = null, CFG = null, BY_ID = {};
  var app = document.getElementById('app');
  var PAGE_SIZE = 12;

  /* ---------- utilidades ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function gs(n) { return 'Gs. ' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function store(key, fb) { try { var v = localStorage.getItem(key); return v ? JSON.parse(v) : fb; } catch (e) { return fb; } }
  function save(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* sin almacenamiento */ } }
  function norm(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
  function plural(n, a, b) { return n + ' ' + (n === 1 ? a : b); }
  function waLink(text) { return 'https://wa.me/' + CFG.whatsapp + (text ? '?text=' + encodeURIComponent(text) : ''); }
  function igLink() { return 'https://www.instagram.com/' + CFG.instagram + '/'; }
  function tkLink() { return 'https://www.tiktok.com/@' + CFG.tiktok; }
  function uniq(arr) { var s = {}; arr.forEach(function (x) { if (x) s[x] = 1; }); return Object.keys(s); }
  function sortEs(arr) { return arr.sort(function (a, b) { return a.localeCompare(b, 'es'); }); }

  var I = {
    heart: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.6A4 4 0 0 1 19 10c0 5.6-7 10-7 10z"/></svg>',
    cmp: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 4v16M17 4v16M3 8h8M13 16h8"/></svg>',
    check: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>',
    x: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    wa: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20l1.3-3.9A8 8 0 1 1 8 19z"/><path d="M9.5 9.5c.3 1.8 2.2 3.7 4 4l1.2-1.2 2 1-.6 1.6c-3.4.4-7.5-3.7-7.1-7.1l1.6-.6 1 2z"/></svg>',
    link: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/></svg>',
    share: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="18" cy="5" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="M8.2 10.8l7.6-4.4M8.2 13.2l7.6 4.4"/></svg>',
    zoom: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4M11 8v6M8 11h6"/></svg>',
    bag: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 8h14l-1.2 11.2a1 1 0 0 1-1 .8H7.2a1 1 0 0 1-1-.8z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/></svg>',
    pin: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 21s-6-5.3-6-11a6 6 0 0 1 12 0c0 5.7-6 11-6 11z"/><circle cx="12" cy="10" r="2.2"/></svg>',
    truck: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 7h11v9H3z"/><path d="M14 10h4l3 3v3h-7"/><circle cx="7" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/></svg>',
    cash: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/></svg>',
    drop: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"/></svg>',
    bottle: '<svg class="bottle" viewBox="0 0 46 96" aria-hidden="true"><rect x="15" y="4" width="16" height="14" rx="2"/><rect x="19" y="18" width="8" height="8"/><path d="M6 38q0-12 17-12t17 12v50a4 4 0 0 1-4 4H10a4 4 0 0 1-4-4z"/></svg>'
  };

  var FAMILIES = {
    Dulce: { d: 'Vainilla, caramelo y frutos secos. Golosa y cercana.', bg: '#FCE8EF' },
    Fresco: { d: 'Cítricos y notas acuáticas. Limpia y liviana.', bg: '#E3F1FA' },
    Amaderado: { d: 'Cedro, sándalo y vetiver. Cálida y elegante.', bg: '#F1E7D6' },
    Especiado: { d: 'Canela, cardamomo y jengibre. Con carácter.', bg: '#FBE7DA' },
    Floral: { d: 'Rosa, jazmín y azahar. Luminosa y suave.', bg: '#F6E6F6' },
    Oriental: { d: 'Ámbar, resinas y vainilla. Envolvente e intensa.', bg: '#EEE8FF' },
    Frutal: { d: 'Mandarina, frutos rojos y tropicales. Jugosa y alegre.', bg: '#FFF1D6' }
  };
  var LEVELS = { Suave: 1, Moderada: 2, Intensa: 3, Corta: 1, Media: 2, Larga: 3 };

  /* ---------- modelo de producto ---------- */
  function withOffer(p, price) { return p.oferta ? Math.round(price * (100 - p.oferta) / 100 / 100) * 100 : price; }
  function variants(p) {
    var out = [];
    if (p.botella) {
      var st = p.botella.stock;
      out.push({ key: 'botella', label: 'Botella' + (p.botella.ml ? ' ' + p.botella.ml + ' ml' : ''), note: 'Original sellada',
        price: withOffer(p, p.botella.precio), old: p.oferta ? p.botella.precio : null,
        state: st == null ? 'ask' : st > 0 ? 'ok' : 'out', max: st == null ? 0 : st });
    }
    p.decants.forEach(function (d) {
      out.push({ key: 'd' + d.ml, ml: d.ml, label: 'Decant ' + d.ml + ' ml', note: 'Porción del original',
        price: withOffer(p, d.precio), old: p.oferta ? d.precio : null,
        state: d.disponible ? 'ok' : 'out', max: Math.floor((p.mlDecant || 0) / d.ml) });
    });
    return out;
  }
  function variant(p, key) { return variants(p).filter(function (v) { return v.key === key; })[0]; }
  function pool(p, mode) {
    var vs = variants(p);
    return mode === 'decants' ? vs.filter(function (v) { return v.key !== 'botella'; }) : vs;
  }
  function priceInfo(p, mode) {
    var vs = pool(p, mode);
    var live = vs.filter(function (v) { return v.state !== 'out'; });
    var use = live.length ? live : vs;
    var min = use.reduce(function (a, v) { return v.price < a.price ? v : a; }, use[0]);
    return { price: min ? min.price : 0, old: min ? min.old : null, many: use.length > 1 };
  }
  function fromPrice(p, mode) { return priceInfo(p, mode).price; }
  function hasDecant(p) { return p.decants.some(function (d) { return d.disponible; }); }
  function avail(p, mode) {
    var vs = pool(p, mode);
    if (vs.some(function (v) { return v.state === 'ok'; })) return { cls: 'ok', txt: 'Disponible' };
    if (vs.some(function (v) { return v.state === 'ask'; })) return { cls: 'ask', txt: 'Consultar disponibilidad' };
    return { cls: 'out', txt: 'Agotado' };
  }
  function presLabel(p, mode) {
    var parts = [];
    if (p.botella && mode !== 'decants') parts.push('Botella');
    var ds = p.decants.filter(function (d) { return d.disponible; }).map(function (d) { return d.ml; });
    if (ds.length) parts.push('Decant ' + ds.join(' / ') + ' ml');
    return parts.join(' · ');
  }
  function img(p, i, lazy) {
    var f = p.fotos && p.fotos[i || 0];
    return f ? '<img src="img/' + esc(f) + '" alt="' + esc(p.marca + ' ' + p.nombre) + '" width="750" height="1000"' + (lazy === false ? '' : ' loading="lazy"') + '>'
      : '<span class="ph">' + I.bottle + '</span>';
  }
  function tags(p) {
    var t = [];
    if (p.oferta) t.push('<span class="tag sale">Oferta −' + p.oferta + '%</span>');
    if (p.novedad) t.push('<span class="tag new">Nuevo</span>');
    if (p.masVendido) t.push('<span class="tag hot">Más vendido</span>');
    if (p.tipo === 'Diseñador') t.push('<span class="tag dark">Diseñador</span>');
    return t.join('');
  }
  function shareUrl(p) {
    return CFG.url ? CFG.url + 'producto/' + p.id + '.html' : location.href.split('#')[0] + '#/p/' + p.id;
  }

  /* ---------- carrito, favoritos, comparador ---------- */
  var cart = store('lp_cart', []);
  var favs = store('lp_favs', []);
  var compare = store('lp_compare', []);
  function cleanState() {
    cart = cart.filter(function (i) { var p = BY_ID[i.id]; var v = p && variant(p, i.v); return v && v.state === 'ok'; });
    favs = favs.filter(function (id) { return BY_ID[id]; });
    compare = compare.filter(function (id) { return BY_ID[id]; }).slice(0, 3);
    save('lp_cart', cart); save('lp_favs', favs); save('lp_compare', compare);
  }
  function units() { return cart.reduce(function (a, i) { return a + i.qty; }, 0); }
  function lineOf(id, key) { return cart.filter(function (i) { return i.id === id && i.v === key; })[0]; }
  function decantMl(id, exceptKey) {
    return cart.reduce(function (a, i) { return (i.id !== id || i.v === 'botella' || i.v === exceptKey) ? a : a + Number(i.v.slice(1)) * i.qty; }, 0);
  }
  function maxQty(p, key) {
    var v = variant(p, key);
    if (!v || v.state !== 'ok') return 0;
    if (key === 'botella') return v.max;
    return Math.floor(((p.mlDecant || 0) - decantMl(p.id, key)) / v.ml);
  }
  function addToCart(id, key, qty) {
    var p = BY_ID[id], line = lineOf(id, key), cur = line ? line.qty : 0;
    var n = Math.min(qty, maxQty(p, key) - cur);
    if (n <= 0) return 0;
    if (line) line.qty += n; else cart.push({ id: id, v: key, qty: n });
    save('lp_cart', cart); updateBadges(true);
    return n;
  }
  function setQty(id, key, qty) {
    var p = BY_ID[id], line = lineOf(id, key);
    if (!line) return;
    line.qty = Math.max(1, Math.min(qty, maxQty(p, key)));
    save('lp_cart', cart); updateBadges();
  }
  function changeVariant(id, from, to) {
    if (from === to) return;
    var p = BY_ID[id], line = lineOf(id, from);
    if (!line) return;
    var qty = line.qty;
    cart = cart.filter(function (i) { return i !== line; });
    var room = maxQty(p, to) - (lineOf(id, to) ? lineOf(id, to).qty : 0);
    var n = Math.max(0, Math.min(qty, room));
    if (n > 0) { var t = lineOf(id, to); if (t) t.qty += n; else cart.push({ id: id, v: to, qty: n }); }
    else cart.push({ id: id, v: from, qty: qty });
    save('lp_cart', cart); updateBadges();
    if (n < qty && n > 0) toast('Ajustamos la cantidad al stock disponible.');
    if (n === 0) toast('No hay stock de esa presentación.');
  }
  function removeLine(id, key) { cart = cart.filter(function (i) { return !(i.id === id && i.v === key); }); save('lp_cart', cart); updateBadges(); }
  function subtotal() { return cart.reduce(function (a, i) { return a + variant(BY_ID[i.id], i.v).price * i.qty; }, 0); }
  function toggleFav(id) { var i = favs.indexOf(id); if (i === -1) favs.push(id); else favs.splice(i, 1); save('lp_favs', favs); updateBadges(); return i === -1; }
  function toggleCompare(id) {
    var i = compare.indexOf(id);
    if (i === -1) { if (compare.length >= 3) { toast('Podés comparar hasta 3 perfumes.'); return false; } compare.push(id); }
    else compare.splice(i, 1);
    save('lp_compare', compare); renderCompareBar();
    return compare.indexOf(id) !== -1;
  }
  function updateBadges(bump) {
    var n = units();
    document.getElementById('cart-count').textContent = n;
    var cb = document.getElementById('cart-btn');
    cb.setAttribute('aria-label', 'Abrir carrito, ' + plural(n, 'producto', 'productos'));
    if (bump) { cb.classList.remove('bump'); void cb.offsetWidth; cb.classList.add('bump'); }
    var fc = document.getElementById('fav-count');
    fc.textContent = favs.length; fc.hidden = !favs.length;
    if (DATA) renderCompareBar();
  }

  /* ---------- tarjeta ---------- */
  function card(p, mode, extra) {
    var fav = favs.indexOf(p.id) !== -1, cmp = compare.indexOf(p.id) !== -1;
    var pi = priceInfo(p, mode), av = avail(p, mode);
    var href = '#/p/' + p.id;
    var btn;
    if (av.cls === 'ok') btn = '<button type="button" class="btn btn-sm btn-block" data-action="quick" data-id="' + p.id + '" data-mode="' + (mode || '') + '">' + I.bag + 'Agregar</button>';
    else if (av.cls === 'ask') btn = '<a class="btn btn-sm btn-block btn-ghost" href="' + waLink('Hola! ¿Tenés disponible ' + p.marca + ' ' + p.nombre + '?') + '" target="_blank" rel="noopener">Consultar</a>';
    else btn = '<button type="button" class="btn btn-sm btn-block" disabled>Agotado</button>';
    var t = tags(p);
    return '<article class="card">' +
      '<div class="card-media"><a href="' + href + '" aria-label="' + esc(p.marca + ' ' + p.nombre) + '">' + img(p) + '</a>' +
        '<div class="card-actions">' +
          '<button type="button" class="round fav" data-action="fav" data-id="' + p.id + '" aria-pressed="' + fav + '" aria-label="' + (fav ? 'Quitar de' : 'Agregar a') + ' favoritos">' + I.heart + '</button>' +
          '<button type="button" class="round cmp" data-action="cmp" data-id="' + p.id + '" aria-pressed="' + cmp + '" aria-label="' + (cmp ? 'Quitar del' : 'Agregar al') + ' comparador">' + I.cmp + '</button>' +
        '</div></div>' +
      '<div class="card-info">' +
        (t ? '<div class="card-tags">' + t + '</div>' : '') +
        '<span class="card-brand">' + esc(p.marca) + '</span>' +
        '<a class="card-name" href="' + href + '">' + esc(p.nombre) + '</a>' +
        '<span class="muted" style="font-size:13px">' + esc(presLabel(p, mode)) + '</span>' +
        '<span class="avail ' + av.cls + '">' + av.txt + '</span>' +
        '<span class="price">' + (pi.many ? '<small>desde</small>' : '') + gs(pi.price) + (pi.old ? '<s>' + gs(pi.old) + '</s>' : '') + '</span>' +
        (extra || '') + btn +
      '</div></article>';
  }

  /* ---------- capas: drawer, sheet, lightbox, toast ----------
     Entran y salen con transiciones (interrumpibles); la salida es más rápida que la entrada. */
  var layer = null, lastFocus = null;
  function show(el) {
    el.hidden = false;
    el.setAttribute('data-state', 'closed');
    void el.offsetWidth;
    el.setAttribute('data-state', 'open');
  }
  function hide(el, instant) {
    el.setAttribute('data-state', 'closed');
    if (instant) { el.hidden = true; return; }
    setTimeout(function () { if (el.getAttribute('data-state') === 'closed') el.hidden = true; }, 220);
  }
  function openLayer(id, html) {
    closeLayer(true);
    var el = document.getElementById(id);
    el.innerHTML = html; show(el);
    if (id !== 'lightbox') show(document.getElementById('overlay'));
    document.body.style.overflow = 'hidden';
    lastFocus = document.activeElement;
    layer = id;
    var f = el.querySelector('[data-autofocus]') || el.querySelector('button, a, input, select');
    if (f) f.focus({ preventScroll: true });
  }
  function closeLayer(silent) {
    if (!layer) return;
    var el = document.getElementById(layer);
    layer = null;
    hide(el, silent);
    var ov = document.getElementById('overlay');
    if (!ov.hidden) hide(ov, silent);
    if (silent) el.innerHTML = '';
    else setTimeout(function () { if (el.hidden) el.innerHTML = ''; }, 240);
    document.body.style.overflow = '';
    if (!silent && lastFocus && document.contains(lastFocus)) lastFocus.focus();
  }
  function toast(msg) {
    var t = document.getElementById('toast');
    t.textContent = msg; show(t);
    clearTimeout(toast.t); toast.t = setTimeout(function () { hide(t); }, 3200);
  }

  function lineHTML(i, ctx) {
    var p = BY_ID[i.id], v = variant(p, i.v), max = maxQty(p, i.v);
    var opts = variants(p).filter(function (x) { return x.state === 'ok'; });
    return '<div class="line">' +
      '<a class="line-media" href="#/p/' + p.id + '" tabindex="-1" aria-hidden="true">' + img(p) + '</a>' +
      '<div class="line-body">' +
        '<div class="line-top"><div><div class="line-brand">' + esc(p.marca) + '</div><a class="line-name" href="#/p/' + p.id + '">' + esc(p.nombre) + '</a></div>' +
        '<button type="button" class="rm" data-action="rm" data-id="' + p.id + '" data-v="' + i.v + '" aria-label="Quitar ' + esc(p.nombre) + '">' + I.x + '</button></div>' +
        (opts.length > 1
          ? '<select data-action="swap" data-id="' + p.id + '" data-v="' + i.v + '" aria-label="Presentación de ' + esc(p.nombre) + '" id="sw-' + ctx + '-' + p.id + '-' + i.v + '">' + opts.map(function (o) {
              return '<option value="' + o.key + '"' + (o.key === i.v ? ' selected' : '') + '>' + o.label + ' (' + gs(o.price) + ')' + '</option>';
            }).join('') + '</select>'
          : '<span class="muted" style="font-size:13px">' + v.label + ', ' + gs(v.price) + '</span>') +
        '<div class="line-bottom"><div class="qty sm"><button type="button" data-action="q-dec" data-id="' + p.id + '" data-v="' + i.v + '" aria-label="Restar uno"' + (i.qty <= 1 ? ' disabled' : '') + '>−</button>' +
          '<output aria-live="polite">' + i.qty + '</output>' +
          '<button type="button" data-action="q-inc" data-id="' + p.id + '" data-v="' + i.v + '" aria-label="Sumar uno"' + (i.qty >= max ? ' disabled' : '') + '>+</button></div>' +
        '<span class="line-price">' + gs(v.price * i.qty) + '</span></div>' +
      '</div></div>';
  }
  function drawerHTML(added) {
    var n = units();
    return '<div class="drawer-head"><h2 id="drawer-title">Tu carrito' + (n ? ' (' + n + ')' : '') + '</h2>' +
      '<button type="button" class="icon-btn" data-action="close" aria-label="Cerrar carrito"' + (added ? '' : ' data-autofocus') + '>' + I.x + '</button></div>' +
      '<div class="drawer-body">' +
        (added ? '<div class="added" role="status">' + I.check + esc(added) + '</div>' : '') +
        (cart.length ? cart.map(function (i) { return lineHTML(i, 'd'); }).join('')
          : '<div class="empty" style="margin-top:12px"><h2>Tu carrito está vacío</h2><p>Elegí una fragancia o probá varias en decant.</p><a class="btn btn-sm" href="#/catalogo" data-action="close-nav">Descubrir perfumes</a></div>') +
      '</div>' +
      (cart.length ? '<div class="drawer-foot"><div class="sum-row sum-total" style="border:0;padding:0"><span>Subtotal</span><span>' + gs(subtotal()) + '</span></div>' +
        '<p class="fine">Envío' + (DATA.cupones.length ? ' y cupones' : '') + ' en el siguiente paso.</p>' +
        '<a class="btn btn-block" href="#/carrito" data-action="close-nav"' + (added ? ' data-autofocus' : '') + '>Finalizar pedido</a>' +
        '<button type="button" class="btn btn-ghost btn-block" data-action="close">Seguir comprando</button></div>' : '');
  }
  function openDrawer(added) { openLayer('drawer', drawerHTML(added)); }
  function refreshDrawer() {
    if (layer !== 'drawer') return;
    var focusId = document.activeElement && document.activeElement.id;
    var act = document.activeElement && document.activeElement.getAttribute('data-action');
    var aid = document.activeElement && document.activeElement.getAttribute('data-id');
    var av = document.activeElement && document.activeElement.getAttribute('data-v');
    document.getElementById('drawer').innerHTML = drawerHTML();
    var el = focusId ? document.getElementById(focusId) : document.querySelector('#drawer [data-action="' + act + '"][data-id="' + aid + '"][data-v="' + av + '"]');
    (el && !el.disabled ? el : document.querySelector('#drawer [data-action="close"]')).focus();
  }

  var quick = null;
  function sheetHTML() {
    var p = BY_ID[quick.id];
    var vs = pool(p, quick.mode);
    var cur = variant(p, quick.v);
    var max = cur && cur.state === 'ok' ? maxQty(p, cur.key) - (lineOf(p.id, cur.key) ? lineOf(p.id, cur.key).qty : 0) : 0;
    if (quick.qty > Math.max(1, max)) quick.qty = Math.max(1, max);
    return '<div class="sheet-head"><h2 id="sheet-title">Elegí la presentación</h2><button type="button" class="icon-btn" data-action="close" aria-label="Cerrar">' + I.x + '</button></div>' +
      '<div class="sheet-body">' +
        '<div class="sheet-prod"><span class="line-media">' + img(p) + '</span><div><div class="line-brand">' + esc(p.marca) + '</div><div class="line-name">' + esc(p.nombre) + '</div></div></div>' +
        '<div class="variants" role="radiogroup" aria-label="Presentación">' + vs.map(function (v) { return variantHTML(v, quick.v, 'qv'); }).join('') + '</div>' +
        (cur && cur.key !== 'botella' ? '<div class="note">Decant: porción del perfume original en un atomizador. No es una botella sellada.</div>' : '') +
        '<div class="buy-row">' + (max > 0
          ? '<div class="qty"><button type="button" data-action="qk-dec" aria-label="Restar uno"' + (quick.qty <= 1 ? ' disabled' : '') + '>−</button><output aria-live="polite">' + quick.qty + '</output><button type="button" data-action="qk-inc" aria-label="Sumar uno"' + (quick.qty >= max ? ' disabled' : '') + '>+</button></div>' +
            '<button type="button" class="btn" data-action="qk-add" data-autofocus>Agregar · ' + gs(cur.price * quick.qty) + '</button>'
          : '<button type="button" class="btn btn-block" disabled>' + (cur && cur.state === 'ok' ? 'Ya tenés todo el stock en el carrito' : 'Elegí una presentación disponible') + '</button>') +
        '</div><a href="#/p/' + p.id + '" data-action="close-nav" class="see-all" style="font-size:14px">Ver ficha completa →</a></div>';
  }
  function openQuick(id, mode) {
    var p = BY_ID[id];
    var ok = pool(p, mode).filter(function (v) { return v.state === 'ok'; });
    if (!ok.length) return;
    if (ok.length === 1 && pool(p, mode).filter(function (v) { return v.state === 'ask'; }).length === 0) {
      var n = addToCart(id, ok[0].key, 1);
      if (n) openDrawer('Agregaste ' + p.nombre + ' · ' + ok[0].label); else toast('No hay más stock de esta presentación.');
      return;
    }
    var pref = ok.filter(function (v) { return v.key !== 'botella'; })[0] || ok[0];
    quick = { id: id, mode: mode, v: pref.key, qty: 1 };
    openLayer('sheet', sheetHTML());
  }
  function refreshSheet(focusSel) {
    document.getElementById('sheet').innerHTML = sheetHTML();
    var el = focusSel && document.querySelector('#sheet ' + focusSel);
    if (el && !el.disabled) el.focus();
  }
  function variantHTML(v, selected, name) {
    var out = v.state === 'out';
    var st = v.state === 'ok' ? '<span class="avail ok">Disponible</span>' : v.state === 'ask' ? '<span class="avail ask">Consultar</span>' : '<span class="avail out">Agotado</span>';
    return '<label class="variant' + (out ? ' is-out' : '') + '"><input type="radio" name="' + name + '" value="' + v.key + '"' + (selected === v.key ? ' checked' : '') + (out ? ' disabled' : '') + '>' +
      '<b>' + v.label + '</b><span class="v-sub">' + v.note + '</span>' +
      '<span class="v-price">' + gs(v.price) + (v.old ? '<s>' + gs(v.old) + '</s>' : '') + '</span>' + st + '</label>';
  }
  function openLightbox(src, alt) {
    openLayer('lightbox', '<button type="button" class="icon-btn close" data-action="close" aria-label="Cerrar foto" data-autofocus>' + I.x + '</button><img src="' + esc(src) + '" alt="' + esc(alt) + '" data-action="close">');
  }

  /* ---------- marco ---------- */
  function renderChrome() {
    var promo = document.getElementById('promo');
    var hasOffers = DATA.productos.some(function (p) { return p.oferta; });
    if (CFG.banner) {
      promo.innerHTML = '<div class="promo">' + esc(CFG.banner) + (hasOffers ? ' · <a href="#/catalogo?destacar=Ofertas">Ver ofertas</a>' : '') + '</div>';
    } else {
      var zonas = CFG.zonas.filter(function (z) { return !z.consultar; }).map(function (z) { return z.zona; });
      promo.innerHTML = '<div class="infobar"><div class="wrap">' +
        '<span>Retiro sin costo en ' + esc(CFG.retiro) + '</span>' +
        (zonas.length ? '<span>Envíos a ' + esc(zonas.join(', ')) + ' y más</span>' : '<span>Envíos con costo según zona</span>') +
        '<span>Pagás con ' + esc(CFG.pagos.join(' o ').toLowerCase()) + '</span></div></div>';
    }
    var social = '<a href="' + igLink() + '" target="_blank" rel="noopener">Instagram @' + esc(CFG.instagram) + '</a>' +
      (CFG.tiktok ? '<a href="' + tkLink() + '" target="_blank" rel="noopener">TikTok @' + esc(CFG.tiktok) + '</a>' : '');
    document.getElementById('footer').innerHTML = '<div class="wrap footer-grid">' +
      '<div class="footer-col"><span class="logo-name">La Plaza</span><span class="muted" style="line-height:1.6">Perfumes originales, botellas y decants. ' + esc(CFG.retiro) + ', Paraguay.</span></div>' +
      '<nav class="footer-col" aria-label="Tienda"><span class="eyebrow">Tienda</span><a href="#/catalogo">Catálogo</a><a href="#/catalogo?genero=Masculino">Masculinos</a><a href="#/catalogo?genero=Femenino">Femeninos</a><a href="#/catalogo?genero=Unisex">Unisex</a><a href="#/decants">Decants</a><a href="#/test">Test olfativo</a><a href="#/comparar">Comparar perfumes</a></nav>' +
      '<nav class="footer-col" aria-label="Ayuda"><span class="eyebrow">Compras</span><a href="#/ayuda?s=envios">Envíos y retiro</a><a href="#/ayuda?s=pagos">Pagos</a><a href="#/ayuda?s=cambios">Cambios y devoluciones</a><a href="#/ayuda?s=terminos">Términos de compra</a><a href="#/ayuda?s=privacidad">Privacidad</a></nav>' +
      '<div class="footer-col"><span class="eyebrow">Contacto</span><a href="' + waLink('Hola! Tengo una consulta.') + '" target="_blank" rel="noopener">WhatsApp +' + esc(CFG.whatsapp) + '</a>' + social + '</div>' +
      '</div><div class="wrap footer-bottom"><span>© ' + new Date().getFullYear() + ' La Plaza Perfumería</span><span>Precios en guaraníes (PYG)</span></div>';
    var wa = document.getElementById('wa-float');
    wa.href = waLink('Hola! Tengo una consulta.'); wa.hidden = false;
  }
  function renderCompareBar() {
    var bar = document.getElementById('compare-bar');
    var onCart = location.hash.indexOf('#/carrito') === 0;
    var n = units();
    if (!cart.length || onCart) { bar.hidden = true; document.body.classList.remove('has-compare'); }
    else {
      var ids = uniq(cart.map(function (i) { return i.id; })).slice(0, 3);
      bar.innerHTML = '<button type="button" class="bar-open" data-action="open-cart" aria-label="Ver carrito">' +
          '<span class="thumbs-mini" aria-hidden="true">' + ids.map(function (id) { return '<span>' + img(BY_ID[id]) + '</span>'; }).join('') + '</span>' +
          '<span class="txt">' + plural(n, 'perfume', 'perfumes') + ' para comprar<small>' + gs(subtotal()) + '</small></span></button>' +
        (compare.length > 1 && location.hash.indexOf('#/comparar') !== 0 ? '<a class="cmp-link" href="#/comparar">Comparar (' + compare.length + ')</a>' : '') +
        '<a class="btn" href="#/carrito">Comprar</a>';
      bar.hidden = false; document.body.classList.add('has-compare');
    }
    document.querySelectorAll('[data-action="cmp"]').forEach(function (b) {
      var on = compare.indexOf(b.getAttribute('data-id')) !== -1;
      b.setAttribute('aria-pressed', on);
      b.setAttribute('aria-label', (on ? 'Quitar del' : 'Agregar al') + ' comparador');
    });
  }
  function setMeta(title, desc) {
    document.title = title;
    var m = document.querySelector('meta[name="description"]');
    if (desc && m) m.setAttribute('content', desc);
  }
  function setJsonLd(obj) {
    var el = document.getElementById('jsonld'); if (el) el.remove();
    if (!obj) return;
    el = document.createElement('script'); el.type = 'application/ld+json'; el.id = 'jsonld';
    el.textContent = JSON.stringify(obj); document.head.appendChild(el);
  }

  /* ---------- inicio ----------
     Portada: escenario negro (las fotos tienen fondo negro y se funden con él), titular cinético,
     marquesina de marcas, índice tipográfico de familias, medidor de decants y pasos. */
  var WORDS = ['aroma.', 'estilo.', 'esencia.'];
  var heroIdx = 0, heroTimer = null, meterMl = null;
  var stagePick = null;
  function shuffle(list) {
    var a = list.slice();
    for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; }
    return a;
  }
  function famPhotos(f, n) {
    var list = DATA.productos.filter(function (p) { return p.familia === f && p.fotos.length; });
    list.sort(function (a, b) { return b.destacado - a.destacado; });
    return list.slice(0, n);
  }
  function decantSizes() {
    var sizes = uniq([].concat.apply([], DATA.productos.map(function (p) { return p.decants.filter(function (d) { return d.disponible; }).map(function (d) { return String(d.ml); }); })))
      .map(Number).sort(function (a, b) { return a - b; });
    return sizes.map(function (ml) {
      var min = Math.min.apply(null, DATA.productos.map(function (p) { var d = p.decants.filter(function (x) { return x.ml === ml && x.disponible; })[0]; return d ? withOffer(p, d.precio) : Infinity; }));
      var n = DATA.productos.filter(function (p) { return p.decants.some(function (d) { return d.ml === ml && d.disponible; }); }).length;
      return { ml: ml, min: min, n: n };
    });
  }
  function viewHome() {
    var P = DATA.productos;
    setMeta('La Plaza Perfumería | Perfumes originales y decants en Paraguay');
    setJsonLd({ '@context': 'https://schema.org', '@type': 'Store', name: 'La Plaza Perfumería', currenciesAccepted: 'PYG',
      address: { '@type': 'PostalAddress', addressLocality: CFG.retiro, addressCountry: 'PY' }, sameAs: [igLink()].concat(CFG.tiktok ? [tkLink()] : []) });
    /* Portada: 3 perfumes al azar entre los que tienen foto y se pueden comprar ya.
       Se sortean una vez por visita, así al volver a Inicio no cambian bajo tus ojos. */
    if (!stagePick) {
      var withPhoto = P.filter(function (p) { return p.fotos.length; });
      var buyable = withPhoto.filter(function (p) { return avail(p).cls === 'ok'; });
      stagePick = shuffle(buyable.length >= 3 ? buyable : withPhoto).slice(0, 3).map(function (p) { return p.id; });
    }
    var stage = stagePick.map(function (id) { return P.filter(function (p) { return p.id === id; })[0]; }).filter(Boolean);
    heroIdx = 0;
    var families = Object.keys(FAMILIES).map(function (f) { return { f: f, n: P.filter(function (p) { return p.familia === f; }).length }; })
      .filter(function (x) { return x.n; }).sort(function (a, b) { return b.n - a.n; });
    var brands = sortEs(uniq(P.map(function (p) { return p.marca; })));
    var sizes = decantSizes();
    var maxMl = sizes.length ? sizes[sizes.length - 1].ml : 10;
    meterMl = sizes.length ? sizes[0].ml : null;
    function strip(title, list, link) {
      if (!list.length) return '';
      return '<section class="wrap section"><div class="section-head"><h2>' + title + '</h2>' + (link ? '<a class="see-all" href="' + link + '">Ver todos</a>' : '') + '</div>' +
        '<div class="grid">' + list.slice(0, 8).map(function (p) { return card(p); }).join('') + '</div></section>';
    }
    var brandRun = brands.map(function (b) { return '<a href="#/catalogo?marca=' + encodeURIComponent(b) + '">' + esc(b) + '</a><span class="mq-sep" aria-hidden="true">/</span>'; }).join('');

    return '' +
    /* 1. Escenario */
    '<section class="wrap stage-wrap" aria-label="Presentación"><div class="stage" id="stage">' +
      '<div class="stage-text">' +
        '<h1 class="sr-only">Tu aroma. Tu estilo. Tu esencia.</h1>' +
        '<p class="kinetic" aria-hidden="true"><span class="k-tu">Tu</span><span class="k-words">' + WORDS.map(function (w, i) { return '<span class="k-word' + (i === 0 ? ' on' : '') + '" data-k="' + i + '">' + w + '</span>'; }).join('') + '</span></p>' +
        '<p class="stage-sub">Perfumes árabes y de diseñador. En botella, o en decant para probar antes de comprar.</p>' +
        '<div class="ctas"><a class="btn btn-light" href="#/catalogo">Descubrir perfumes</a><a class="btn btn-lilac" href="#/test">Encontrar mi fragancia</a></div>' +
        (stage.length > 1 ? '<div class="k-tabs" role="tablist" aria-label="Perfumes para descubrir">' + stage.map(function (p, i) {
          return '<button type="button" role="tab" class="k-tab' + (i === 0 ? ' on' : '') + '" data-action="hero" data-k="' + i + '" aria-selected="' + (i === 0) + '" aria-controls="kp-' + i + '">' + esc(p.nombre) + '</button>';
        }).join('') + '</div>' : '') +
      '</div>' +
      '<div class="stage-prods">' + stage.map(function (p, i) {
        var pi = priceInfo(p);
        return '<div class="k-prod' + (i === 0 ? ' on' : '') + '" data-k="' + i + '" id="kp-' + i + '" role="tabpanel">' +
          '<a class="k-photo" href="#/p/' + p.id + '" aria-label="' + esc(p.marca + ' ' + p.nombre) + '">' + img(p, 0, false) + '</a>' +
          '<div class="k-cap"><div><span class="k-brand">' + esc(p.marca) + '</span><a class="k-name" href="#/p/' + p.id + '">' + esc(p.nombre) + '</a></div>' +
          '<div class="k-buy"><span class="k-price">' + (pi.many ? 'desde ' : '') + gs(pi.price) + '</span>' +
          (avail(p).cls === 'ok' ? '<button type="button" class="btn btn-sm btn-light" data-action="quick" data-id="' + p.id + '">' + I.bag + 'Agregar</button>' : '') + '</div></div></div>';
      }).join('') + '</div>' +
    '</div></section>' +

    /* 2. Marquesina de marcas */
    (brands.length > 2 ? '<section class="marquee" aria-label="Marcas"><div class="mq-track"><div class="mq-run">' + brandRun + '</div><div class="mq-run" aria-hidden="true">' + brandRun.replace(/<a /g, '<a tabindex="-1" ') + '</div></div></section>' : '') +

    /* 3. Índice de familias */
    '<section class="wrap section fam-index"><h2 class="fi-title">¿Cuál es tu estilo?</h2>' +
      '<div class="fi-grid"><ol class="fi-list">' + families.map(function (x, i) {
        var thumbs = famPhotos(x.f, 3);
        return '<li><a class="fi-row' + (i === 0 ? ' on' : '') + '" data-fam="' + x.f + '" href="#/catalogo?familia=' + encodeURIComponent(x.f) + '">' +
          '<span class="fi-name">' + x.f + '</span><span class="fi-desc">' + FAMILIES[x.f].d + '</span>' +
          '<span class="fi-thumbs" aria-hidden="true">' + thumbs.map(function (p) { return '<span>' + img(p) + '</span>'; }).join('') + '</span>' +
          '<span class="fi-count">' + plural(x.n, 'perfume', 'perfumes') + '</span></a></li>';
      }).join('') +
      '<li><a class="fi-row fi-test" href="#/test"><span class="fi-name">¿No sabés?</span><span class="fi-desc">Respondé 6 preguntas y te recomendamos.</span><span class="fi-count">Encontrar mi fragancia</span></a></li></ol>' +
      '<div class="fi-preview" id="fi-preview" aria-hidden="true">' + families.map(function (x, i) {
        return '<div class="fi-set' + (i === 0 ? ' on' : '') + '" data-fam="' + x.f + '">' + famPhotos(x.f, 3).map(function (p) { return '<span>' + img(p) + '</span>'; }).join('') + '</div>';
      }).join('') + '</div></div></section>' +

    strip('Más vendidos', P.filter(function (p) { return p.masVendido; }), '#/catalogo?destacar=Más vendidos') +
    strip('Ofertas', P.filter(function (p) { return p.oferta; }), '#/catalogo?destacar=Ofertas') +
    strip('Novedades', P.filter(function (p) { return p.novedad; }), '#/catalogo?destacar=Nuevos') +

    /* 4. Medidor de decants */
    (sizes.length ? '<section class="band"><div class="wrap section meter-sec">' +
      '<div class="ms-text"><h2>¿Cuánto rinde un decant?</h2>' +
        '<p>Un decant es una porción del perfume original en un atomizador. Elegí la medida y mirá para cuánto te alcanza.</p>' +
        '<div class="ms-sizes" role="radiogroup" aria-label="Medida">' + sizes.map(function (s) {
          return '<button type="button" role="radio" class="ms-size" data-action="ml" data-ml="' + s.ml + '" aria-checked="' + (s.ml === meterMl) + '">' + s.ml + ' ml</button>';
        }).join('') + '</div>' +
        '<dl class="ms-stats" id="ms-stats"></dl>' +
        '<div class="ctas"><a class="btn" href="#/decants" id="ms-cta">Ver decants</a></div>' +
        '<p class="fine">Cálculo aproximado: una atomización equivale a unos 0,1 ml y contamos 3 atomizaciones por día.</p></div>' +
      '<div class="ms-visual" aria-hidden="true"><span class="ms-big" id="ms-big"></span>' +
        '<div class="vial"><span class="vial-cap"></span><span class="vial-neck"></span><span class="vial-glass"><span class="vial-liquid" id="vial-liquid"></span>' +
        sizes.map(function (s) { return '<span class="vial-tick" style="bottom:' + (s.ml / maxMl * 100) + '%"><i>' + s.ml + ' ml</i></span>'; }).join('') +
        '</span></div></div>' +
    '</div></section>' : '') +

    /* 5. Pasos como manifiesto */
    '<section class="wrap section steps"><h2 class="sr-only">Cómo comprar</h2><div class="steps-row">' +
      '<div class="step"><span class="verb">Elegí.</span><p>Botella original o decant. Filtrá por familia, ocasión o precio.</p></div>' +
      '<div class="step"><span class="verb">Armá.</span><p>Sumá al carrito y elegí envío o retiro sin costo en ' + esc(CFG.retiro) + '.</p></div>' +
      '<div class="step"><span class="verb">Enviá.</span><p>Te llega el pedido escrito por WhatsApp y coordinamos el pago con ' + esc(CFG.pagos.join(' o ').toLowerCase()) + '.</p></div>' +
      '</div>' +
      (CFG.zonas.length ? '<ul class="zones" aria-label="Costos de envío">' + CFG.zonas.map(function (z) {
        return '<li><span>' + esc(z.zona) + '</span><b>' + (z.consultar ? 'consultar' : z.costo === 0 ? 'sin costo' : gs(z.costo)) + '</b></li>';
      }).join('') + '<li><span>Retiro</span><b>sin costo</b></li></ul>' : '') +
    '</section>' +

    /* 6. Redes */
    '<section class="wrap section social2"><p>Nuevos ingresos y promos, primero en Instagram</p>' +
      '<a class="handle" href="' + igLink() + '" target="_blank" rel="noopener">@' + esc(CFG.instagram) + '</a>' +
      (CFG.tiktok ? '<a class="handle sm" href="' + tkLink() + '" target="_blank" rel="noopener">TikTok @' + esc(CFG.tiktok) + '</a>' : '') +
    '</section>';
  }
  function setHero(i) {
    var root = document.getElementById('stage');
    if (!root) return;
    heroIdx = i;
    root.querySelectorAll('[data-k]').forEach(function (el) { el.classList.toggle('on', Number(el.getAttribute('data-k')) === i); });
    root.querySelectorAll('.k-tab').forEach(function (t) { t.setAttribute('aria-selected', Number(t.getAttribute('data-k')) === i); });
  }
  function startHero() {
    clearInterval(heroTimer); heroTimer = null;
    var root = document.getElementById('stage');
    if (!root || root.querySelectorAll('.k-prod').length < 2) return;
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    heroTimer = setInterval(function () {
      if (!document.getElementById('stage')) { clearInterval(heroTimer); return; }
      // Pausa si la pestaña no está visible, si el mouse está encima o si hay foco adentro.
      if (document.hidden || root.matches(':hover') || root.contains(document.activeElement)) return;
      setHero((heroIdx + 1) % root.querySelectorAll('.k-prod').length);
    }, 3800);
  }
  function setMeter(ml) {
    var sizes = decantSizes();
    var s = sizes.filter(function (x) { return x.ml === ml; })[0];
    if (!s) return;
    meterMl = ml;
    var maxMl = sizes[sizes.length - 1].ml;
    var sprays = ml * 10, days = Math.floor(sprays / 3);
    document.getElementById('ms-big').innerHTML = ml + '<small>ml</small>';
    document.getElementById('vial-liquid').style.clipPath = 'inset(' + (100 - ml / maxMl * 100) + '% 0 0 0)';
    document.getElementById('ms-stats').innerHTML =
      '<div><dt>Rinde</dt><dd>unas ' + sprays + ' atomizaciones</dd></div>' +
      '<div><dt>Te dura</dt><dd>unos ' + days + ' días</dd></div>' +
      '<div><dt>Precio</dt><dd>desde ' + gs(s.min) + '</dd></div>' +
      '<div><dt>Disponible en</dt><dd>' + plural(s.n, 'perfume', 'perfumes') + '</dd></div>';
    document.getElementById('ms-cta').setAttribute('href', '#/decants?tamano=' + encodeURIComponent(ml + ' ml'));
    document.querySelectorAll('.ms-size').forEach(function (b) { b.setAttribute('aria-checked', Number(b.getAttribute('data-ml')) === ml); });
  }
  function setFam(f) {
    var pv = document.getElementById('fi-preview');
    if (!pv) return;
    pv.querySelectorAll('.fi-set').forEach(function (x) { x.classList.toggle('on', x.getAttribute('data-fam') === f); });
    document.querySelectorAll('.fi-row[data-fam]').forEach(function (x) { x.classList.toggle('on', x.getAttribute('data-fam') === f); });
  }
  function afterHome() {
    startHero();
    if (meterMl != null) setMeter(meterMl);
  }

  /* ---------- catálogo ---------- */
  var PRICE_BUCKETS = [
    { id: 'hasta50', label: 'Hasta Gs. 50.000', min: 0, max: 50000 },
    { id: '50a200', label: 'Gs. 50.000 a 200.000', min: 50001, max: 200000 },
    { id: '200a300', label: 'Gs. 200.000 a 300.000', min: 200001, max: 300000 },
    { id: 'mas300', label: 'Más de Gs. 300.000', min: 300001, max: Infinity }
  ];
  function filterDefs(mode) {
    var P = DATA.productos;
    var defs = [
      { key: 'genero', title: 'Género', vals: sortEs(uniq(P.map(function (p) { return p.genero; }))), test: function (p, v) { return p.genero === v; } },
      { key: 'marca', title: 'Marca', vals: sortEs(uniq(P.map(function (p) { return p.marca; }))), test: function (p, v) { return p.marca === v; } },
      { key: 'pres', title: 'Presentación', vals: mode === 'decants' ? [] : ['Botella', 'Decant'], test: function (p, v) { return v === 'Botella' ? !!p.botella : hasDecant(p); } },
      { key: 'tamano', title: 'Tamaño', vals: mode === 'decants' ? uniq([].concat.apply([], P.map(function (p) { return p.decants.filter(function (d) { return d.disponible; }).map(function (d) { return d.ml + ' ml'; }); }))).sort(function (a, b) { return parseInt(a) - parseInt(b); }) : [],
        test: function (p, v) { return p.decants.some(function (d) { return d.disponible && d.ml + ' ml' === v; }); } },
      { key: 'precio', title: 'Precio', vals: PRICE_BUCKETS.map(function (b) { return b.id; }), label: function (v) { return PRICE_BUCKETS.filter(function (b) { return b.id === v; })[0].label; },
        test: function (p, v) { var b = PRICE_BUCKETS.filter(function (x) { return x.id === v; })[0]; return pool(p, mode).some(function (x) { return x.state !== 'out' && x.price >= b.min && x.price <= b.max; }); } },
      { key: 'familia', title: 'Familia olfativa', vals: sortEs(uniq(P.map(function (p) { return p.familia; }))), test: function (p, v) { return p.familia === v; } },
      { key: 'ocasion', title: 'Ocasión', vals: sortEs(uniq([].concat.apply([], P.map(function (p) { return p.ocasion; })))), test: function (p, v) { return p.ocasion.indexOf(v) !== -1; } },
      { key: 'momento', title: 'Momento', vals: ['Día', 'Noche'].filter(function (v) { return P.some(function (p) { return p.momento.indexOf(v) !== -1; }); }), test: function (p, v) { return p.momento === v || p.momento === 'Día y noche'; } },
      { key: 'clima', title: 'Clima', vals: ['Calor', 'Frío'].filter(function (v) { return P.some(function (p) { return p.clima === v; }); }), test: function (p, v) { return p.clima === v || p.clima === 'Todo el año'; } },
      { key: 'intensidad', title: 'Intensidad', vals: ['Suave', 'Moderada', 'Intensa'].filter(function (v) { return P.some(function (p) { return p.intensidad === v; }); }), test: function (p, v) { return p.intensidad === v; } },
      { key: 'duracion', title: 'Duración', vals: ['Corta', 'Media', 'Larga'].filter(function (v) { return P.some(function (p) { return p.duracion === v; }); }), test: function (p, v) { return p.duracion === v; } },
      { key: 'destacar', title: 'Destacados', vals: [['Ofertas', 'oferta'], ['Nuevos', 'novedad'], ['Más vendidos', 'masVendido']].filter(function (x) { return P.some(function (p) { return p[x[1]]; }); }).map(function (x) { return x[0]; }),
        test: function (p, v) { return v === 'Ofertas' ? !!p.oferta : v === 'Nuevos' ? p.novedad : p.masVendido; } }
    ];
    return defs.filter(function (d) { return d.vals.length; });
  }
  var cat = null;
  function parseQuery(qs, mode) {
    var st = { q: '', sort: 'rel', fav: false, limit: PAGE_SIZE, mode: mode || '' };
    filterDefs(mode).forEach(function (f) { st[f.key] = []; });
    new URLSearchParams(qs || '').forEach(function (v, k) {
      if (k === 'q') st.q = v; else if (k === 'sort') st.sort = v; else if (k === 'fav') st.fav = v === '1';
      else if (st[k]) st[k] = v.split(',').filter(Boolean);
    });
    return st;
  }
  function queryString(st) {
    var p = new URLSearchParams();
    if (st.q) p.set('q', st.q);
    filterDefs(st.mode).forEach(function (f) { if (st[f.key] && st[f.key].length) p.set(f.key, st[f.key].join(',')); });
    if (st.sort !== 'rel') p.set('sort', st.sort);
    if (st.fav) p.set('fav', '1');
    var s = p.toString();
    return s ? '?' + s : '';
  }
  function matches(p, st, skip) {
    if (st.mode === 'decants' && !hasDecant(p)) return false;
    if (st.fav && favs.indexOf(p.id) === -1) return false;
    if (st.q) {
      var hay = norm([p.marca, p.nombre, p.familia, p.notas.salida, p.notas.corazon, p.notas.fondo, p.ocasion.join(' ')].join(' '));
      if (!norm(st.q).split(/\s+/).every(function (w) { return hay.indexOf(w) !== -1; })) return false;
    }
    return filterDefs(st.mode).every(function (f) {
      if (f.key === skip || !st[f.key] || !st[f.key].length) return true;
      return st[f.key].some(function (v) { return f.test(p, v); });
    });
  }
  function viewCatalog(qs, mode) {
    cat = parseQuery(qs, mode);
    var dec = mode === 'decants';
    var title = cat.fav ? 'Tus favoritos' : dec ? 'Decants' : cat.genero && cat.genero.length === 1 ? 'Perfumes ' + cat.genero[0].toLowerCase() + (cat.genero[0] === 'Unisex' ? '' : 's') : 'Catálogo';
    setMeta(title + ' | La Plaza Perfumería', dec ? 'Decants de perfumes originales para probar antes de comprar la botella.' : 'Perfumes originales árabes y de diseñador en Paraguay.');
    setJsonLd(null);
    var wide = window.matchMedia('(min-width: 641px)').matches;
    var sorts = [['rel', 'Relevancia'], ['precio-asc', 'Precio: menor a mayor'], ['precio-desc', 'Precio: mayor a menor'], ['nombre', 'Nombre A-Z']];
    if (DATA.productos.some(function (p) { return p.novedad; })) sorts.splice(1, 0, ['nuevos', 'Novedades']);
    return '<div class="wrap page">' +
      '<nav class="crumbs" aria-label="Ruta"><a href="#/">Inicio</a> / ' + (dec ? 'Decants' : 'Catálogo') + '</nav>' +
      (dec ? '<div class="decant-hero"><div><h1>Decants</h1><p>Porciones del perfume original en atomizador, para probar con calma. No son botellas selladas. Podés sumar varios al mismo pedido y armar tu propio kit.</p></div><a class="btn btn-lilac" href="#/test">Encontrar mi fragancia</a></div>'
        : '<h1 class="page-title">' + esc(title) + '</h1><p class="page-intro">Buscá por nombre, marca o nota: por ejemplo «vainilla», «Lattafa» o «cítricos».</p>') +
      '<div class="toolbar">' +
        '<label class="field-box q"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4"/></svg>' +
          '<input type="search" id="cat-q" aria-label="Buscar por perfume, marca o nota" placeholder="Buscar por perfume, marca o nota" value="' + esc(cat.q) + '"></label>' +
        '<label class="field-box sort">Ordenar<select id="cat-sort">' + sorts.map(function (o) { return '<option value="' + o[0] + '"' + (cat.sort === o[0] ? ' selected' : '') + '>' + o[1] + '</option>'; }).join('') + '</select></label>' +
      '</div>' +
      '<div class="cat-layout">' +
        '<details class="filters" id="filters"' + (wide ? ' open' : '') + '><summary>Filtros <span id="f-count"></span></summary><div class="filter-groups" id="filter-groups"></div></details>' +
        '<section class="results" aria-label="Resultados"><div class="chips" id="chips"></div><div id="results"></div></section>' +
      '</div></div>';
  }
  function renderCatalogParts(focusId) {
    var st = cat, defs = filterDefs(st.mode);
    document.getElementById('filter-groups').innerHTML = defs.map(function (f) {
      return '<fieldset><legend>' + f.title + '</legend>' + f.vals.map(function (v) {
        var n = DATA.productos.filter(function (p) { return matches(p, st, f.key) && f.test(p, v); }).length;
        var on = st[f.key].indexOf(v) !== -1;
        var id = 'f-' + f.key + '-' + norm(v).replace(/[^a-z0-9]+/g, '-');
        return '<label class="check' + (n === 0 && !on ? ' is-empty' : '') + '"><input type="checkbox" id="' + id + '" data-filter="' + f.key + '" value="' + esc(v) + '"' + (on ? ' checked' : '') + '><span class="box">' + I.check + '</span>' + esc(f.label ? f.label(v) : v) + '<span class="n">' + n + '</span></label>';
      }).join('') + '</fieldset>';
    }).join('');

    var list = DATA.productos.filter(function (p) { return matches(p, st); });
    var fp = function (p) { return fromPrice(p, st.mode); };
    if (st.sort === 'precio-asc') list.sort(function (a, b) { return fp(a) - fp(b); });
    else if (st.sort === 'precio-desc') list.sort(function (a, b) { return fp(b) - fp(a); });
    else if (st.sort === 'nombre') list.sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });
    else if (st.sort === 'nuevos') list.sort(function (a, b) { return (b.novedad - a.novedad) || (b.orden - a.orden); });
    else list.sort(function (a, b) {
      var sa = avail(a, st.mode).cls === 'ok' ? 1 : 0, sb = avail(b, st.mode).cls === 'ok' ? 1 : 0;
      return (sb - sa) || ((b.masVendido + b.destacado + b.novedad) - (a.masVendido + a.destacado + a.novedad)) || a.orden - b.orden;
    });

    var chips = [];
    defs.forEach(function (f) { st[f.key].forEach(function (v) { chips.push({ key: f.key, v: v, label: f.label ? f.label(v) : v }); }); });
    if (st.q) chips.push({ key: 'q', v: st.q, label: '«' + st.q + '»' });
    if (st.fav) chips.push({ key: 'fav', v: '1', label: 'Favoritos' });
    document.getElementById('f-count').textContent = chips.length ? '(' + chips.length + ')' : '';
    document.getElementById('chips').innerHTML = '<span class="count-label" role="status">' + plural(list.length, 'fragancia', 'fragancias') + '</span>' +
      chips.map(function (c) { return '<button type="button" class="chip" data-action="unchip" data-key="' + c.key + '" data-v="' + esc(c.v) + '" aria-label="Quitar filtro ' + esc(c.label) + '">' + esc(c.label) + '<span aria-hidden="true">×</span></button>'; }).join('') +
      (chips.length ? '<button type="button" class="link-btn" data-action="clear">Limpiar filtros</button>' : '');

    var shown = list.slice(0, st.limit);
    document.getElementById('results').innerHTML = list.length
      ? '<div class="grid">' + shown.map(function (p) { return card(p, st.mode); }).join('') + '</div>' +
        (list.length > shown.length ? '<div class="more"><button type="button" class="btn btn-ghost" data-action="more">Ver más (' + (list.length - shown.length) + ')</button></div>' : '')
      : '<div class="empty"><h2>' + (st.fav && !favs.length ? 'Todavía no guardaste favoritos' : 'No encontramos perfumes') + '</h2>' +
        '<p>' + (st.fav && !favs.length ? 'Tocá el corazón de un perfume para guardarlo acá.' : 'Probá quitando algún filtro o pedinos una recomendación.') + '</p>' +
        '<div class="row">' + (chips.length ? '<button type="button" class="btn btn-ghost btn-sm" data-action="clear">Limpiar filtros</button>' : '') +
        '<a class="btn btn-sm" href="' + waLink('Hola! Busco una recomendación de perfume.') + '" target="_blank" rel="noopener">Pedir recomendación</a></div></div>';
    if (focusId) { var el = document.getElementById(focusId); if (el) el.focus(); }
  }
  function syncCatalog(focusId) {
    history.replaceState(null, '', '#/' + (cat.mode === 'decants' ? 'decants' : 'catalogo') + queryString(cat));
    renderCatalogParts(focusId);
  }

  /* ---------- ficha ---------- */
  var pdp = null;
  function viewProduct(id) {
    var p = BY_ID[id];
    if (!p) { pdp = null; return viewNotFound(); }
    var vs = variants(p);
    var first = vs.filter(function (v) { return v.state === 'ok' && v.key !== 'botella'; })[0] || vs.filter(function (v) { return v.state === 'ok'; })[0] || vs.filter(function (v) { return v.state === 'ask'; })[0] || vs[0];
    pdp = { id: id, v: first ? first.key : null, qty: 1, foto: 0 };
    var title = p.marca + ' ' + p.nombre;
    setMeta(title + ' | La Plaza Perfumería', (p.descripcion ? p.descripcion + ' ' : '') + 'Desde ' + gs(fromPrice(p)) + '.');
    setJsonLd({ '@context': 'https://schema.org', '@type': 'Product', name: p.nombre, brand: { '@type': 'Brand', name: p.marca }, description: p.descripcion,
      image: p.fotos.map(function (f) { return 'img/' + f; }),
      offers: vs.map(function (v) { return { '@type': 'Offer', name: v.label, price: v.price, priceCurrency: 'PYG', availability: 'https://schema.org/' + (v.state === 'ok' ? 'InStock' : v.state === 'out' ? 'OutOfStock' : 'LimitedAvailability') }; }) });
    var similar = DATA.productos.filter(function (x) { return x.id !== p.id && x.familia === p.familia; })
      .concat(DATA.productos.filter(function (x) { return x.id !== p.id && x.familia !== p.familia && x.genero === p.genero; })).slice(0, 4);
    var n = p.notas;
    var when = [];
    if (p.momento) p.momento.split(' y ').forEach(function (m) { when.push('<span>' + (m === 'Día' || m === 'día' ? 'De día' : m.toLowerCase() === 'noche' ? 'De noche' : m) + '</span>'); });
    if (p.clima) when.push('<span>' + (p.clima === 'Todo el año' ? 'Todo el año' : 'Clima ' + p.clima.toLowerCase()) + '</span>');
    p.ocasion.forEach(function (o) { when.push('<span>' + esc(o) + '</span>'); });
    var meters = [['Intensidad', p.intensidad], ['Duración', p.duracion]].filter(function (m) { return m[1]; });
    return '<div class="wrap page">' +
      '<nav class="crumbs" aria-label="Ruta"><a href="#/">Inicio</a> / <a href="#/catalogo">Catálogo</a> / ' + esc(p.nombre) + '</nav>' +
      '<div class="pdp">' +
        '<div class="gallery" id="gallery"></div>' +
        '<div class="pdp-info">' +
          (tags(p) ? '<div class="pdp-tags">' + tags(p).replace(/tag dark/g, 'tag') + '</div>' : '') +
          '<div><div class="pdp-brand">' + esc(p.marca) + '</div><h1>' + esc(p.nombre) + '</h1>' +
          '<div class="pdp-meta">' + [p.concentracion, p.genero, p.familia ? 'familia ' + p.familia.toLowerCase() : '', p.tipo].filter(Boolean).map(esc).join(', ') + '</div></div>' +
          (p.descripcion ? '<p class="pdp-desc">' + esc(p.descripcion) + '</p>' : '') +
          '<div class="pdp-price" id="pdp-price"></div>' +
          '<div><div class="label" id="pres-l">Presentación</div><div class="variants" role="radiogroup" aria-labelledby="pres-l" id="variants"></div></div>' +
          '<div id="pdp-note"></div>' +
          '<div class="buy-row" id="buy-row"></div>' +
          '<div class="share-row" id="share-row"></div>' +
          '<div class="blocks">' +
            ((n.salida || n.corazon || n.fondo) ? '<div class="block block-wide"><div class="label">Pirámide olfativa</div><div class="pyramid">' + [['Salida', n.salida, 'Lo primero que sentís'], ['Corazón', n.corazon, 'A los pocos minutos'], ['Fondo', n.fondo, 'Lo que queda horas después']].filter(function (x) { return x[1]; }).map(function (x, i) { return '<div class="tier t' + i + '"><span class="tier-k"><b>' + x[0] + '</b><small>' + x[2] + '</small></span><span class="tier-v">' + esc(x[1]) + '</span></div>'; }).join('') + '</div></div>' : '') +
            (when.length ? '<div class="block"><div class="label">Cuándo usarlo</div><div class="when">' + when.join('') + '</div></div>' : '') +
            (meters.length ? '<div class="block"><div class="label">Rendimiento</div>' + meters.map(function (m) { return '<div class="meter"><span>' + m[0] + '</span><b>' + esc(m[1]) + '</b></div>'; }).join('') + '<p class="fine">Estimado: depende de tu piel y del clima.</p></div>' : '') +
            '<div class="block"><div class="label">Entrega y pago</div><div class="info-list"><span>Retiro sin costo en ' + esc(CFG.retiro) + '.</span><span>Envíos: ' + esc(CFG.zonas.length ? CFG.zonas.map(function (z) { return z.zona + (z.consultar ? ' (consultar)' : z.costo === 0 ? ' sin costo' : ' ' + gs(z.costo)); }).join(', ') : 'según zona') + '.</span><span>Pago: ' + esc(CFG.pagos.join(' o ').toLowerCase()) + '.</span>' +
              '<a href="#" id="ask-link" target="_blank" rel="noopener" style="font-weight:700">' + 'Consultar por WhatsApp →</a></div></div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      (similar.length ? '<section class="section" style="padding-bottom:0"><div class="section-head"><h2>Perfumes similares</h2></div><div class="grid">' + similar.map(function (x) { return card(x); }).join('') + '</div></section>' : '') +
    '</div>';
  }
  function renderGallery() {
    var p = BY_ID[pdp.id];
    var f = p.fotos;
    document.getElementById('gallery').innerHTML =
      (f.length ? '<button type="button" class="gallery-main" data-action="zoom" aria-label="Ampliar foto de ' + esc(p.nombre) + '">' + img(p, pdp.foto, false) + '<span class="zoom-hint">' + I.zoom + 'Ampliar</span></button>'
        : '<div class="gallery-main" style="cursor:default">' + img(p) + '</div>') +
      (f.length > 1 ? '<div class="thumbs">' + f.map(function (_, i) { return '<button type="button" data-action="thumb" data-i="' + i + '" aria-label="Foto ' + (i + 1) + '" aria-current="' + (i === pdp.foto) + '">' + img(p, i) + '</button>'; }).join('') + '</div>' : '');
  }
  function renderPdpParts(focusKey) {
    var p = BY_ID[pdp.id], cur = variant(p, pdp.v);
    document.getElementById('variants').innerHTML = variants(p).map(function (v) { return variantHTML(v, pdp.v, 'variant'); }).join('');
    document.getElementById('pdp-price').innerHTML = cur ? gs(cur.price * (cur.state === 'ok' ? pdp.qty : 1)) + (cur.old ? '<s>' + gs(cur.old * (cur.state === 'ok' ? pdp.qty : 1)) + '</s>' : '') + '<small>' + cur.label + (pdp.qty > 1 && cur.state === 'ok' ? ' × ' + pdp.qty : '') + '</small>' : '';
    document.getElementById('pdp-note').innerHTML = cur && cur.key !== 'botella' ? '<div class="note">Decant: porción del perfume original en un atomizador. No es una botella sellada.</div>' : '';
    var fav = favs.indexOf(p.id) !== -1, cmp = compare.indexOf(p.id) !== -1;
    var extras = '<button type="button" class="square fav" data-action="fav" data-id="' + p.id + '" aria-pressed="' + fav + '" aria-label="' + (fav ? 'Quitar de' : 'Agregar a') + ' favoritos">' + I.heart + '</button>' +
      '<button type="button" class="square" data-action="cmp" data-id="' + p.id + '" aria-pressed="' + cmp + '" aria-label="' + (cmp ? 'Quitar del' : 'Agregar al') + ' comparador">' + I.cmp + '</button>';
    var buy;
    if (!cur || cur.state === 'out') buy = '<button class="btn" type="button" disabled>Agotado</button>' + extras;
    else if (cur.state === 'ask') buy = '<a class="btn" href="' + waLink('Hola! ¿Tenés disponible ' + p.marca + ' ' + p.nombre + ' (' + cur.label + ')?') + '" target="_blank" rel="noopener">' + I.wa + 'Consultar disponibilidad</a>' + extras;
    else {
      var max = maxQty(p, cur.key) - (lineOf(p.id, cur.key) ? lineOf(p.id, cur.key).qty : 0);
      if (pdp.qty > Math.max(1, max)) pdp.qty = Math.max(1, max);
      buy = max <= 0 ? '<button type="button" class="btn btn-ghost" data-action="open-cart">Ya tenés todo el stock en el carrito</button>' + extras
        : '<div class="qty"><button type="button" data-action="pdp-dec" aria-label="Restar uno"' + (pdp.qty <= 1 ? ' disabled' : '') + '>−</button><output aria-live="polite">' + pdp.qty + '</output><button type="button" data-action="pdp-inc" aria-label="Sumar uno"' + (pdp.qty >= max ? ' disabled' : '') + '>+</button></div>' +
          '<button class="btn" type="button" data-action="add">' + I.bag + 'Agregar al carrito</button>' + extras;
    }
    document.getElementById('buy-row').innerHTML = buy;
    var url = shareUrl(p);
    document.getElementById('share-row').innerHTML = '<span>Compartir:</span>' +
      '<a href="https://wa.me/?text=' + encodeURIComponent('Mirá este perfume: ' + p.marca + ' ' + p.nombre + ' ' + url) + '" target="_blank" rel="noopener">' + I.wa + 'WhatsApp</a>' +
      '<a href="https://www.facebook.com/sharer/sharer.php?u=' + encodeURIComponent(url) + '" target="_blank" rel="noopener">Facebook</a>' +
      '<button type="button" data-action="copy">' + I.link + 'Copiar enlace</button>' +
      (navigator.share ? '<button type="button" data-action="native-share">' + I.share + 'Más</button>' : '');
    document.getElementById('ask-link').href = waLink('Hola! Quiero consultar por ' + p.marca + ' ' + p.nombre + (cur ? ' (' + cur.label + ')' : '') + '.');
    if (focusKey) { var r = document.querySelector('#variants input[value="' + focusKey + '"]'); if (r) r.focus(); }
  }

  /* ---------- carrito ---------- */
  var checkout = store('lp_checkout', { name: '', phone: '', mode: 'envio', zona: '', addr: '', pago: '', notes: '', cupon: '' });
  var tried = false, sent = false, couponMsg = '';
  function coupon() {
    if (!checkout.cupon) return null;
    var c = DATA.cupones.filter(function (x) { return x.codigo === checkout.cupon; })[0];
    if (!c) return null;
    var sub = subtotal();
    if (sub < c.minimo) return { c: c, amount: 0, msg: 'Este cupón vale para compras desde ' + gs(c.minimo) + '.' };
    var amount = c.tipo === 'monto' ? Math.min(c.valor, sub) : Math.round(sub * c.valor / 100 / 100) * 100;
    return { c: c, amount: amount, msg: '' };
  }
  function shipping() {
    if (checkout.mode !== 'envio') return { cost: 0, known: true, label: 'Retiro en ' + CFG.retiro };
    if (!CFG.zonas.length) return { cost: 0, known: false, label: 'Envío (según zona)' };
    var z = CFG.zonas.filter(function (x) { return x.zona === checkout.zona; })[0];
    if (!z) return { cost: 0, known: false, label: 'Envío (elegí la zona)' };
    if (z.consultar) return { cost: 0, known: false, label: 'Envío a ' + z.zona + ' (a confirmar)' };
    return { cost: z.costo, known: true, label: 'Envío a ' + z.zona };
  }
  function totals() {
    var sub = subtotal(), cp = coupon(), sh = shipping();
    var disc = cp ? cp.amount : 0;
    var saved = cart.reduce(function (a, i) { var v = variant(BY_ID[i.id], i.v); return a + (v.old ? (v.old - v.price) * i.qty : 0); }, 0);
    return { sub: sub, disc: disc, cp: cp, sh: sh, saved: saved, total: sub - disc + sh.cost };
  }
  function errors() {
    var e = {};
    if (checkout.name.trim().length < 2) e.name = 'Escribí tu nombre.';
    if (checkout.phone.replace(/\D/g, '').length < 9) e.phone = 'Escribí un teléfono válido, por ejemplo 0981 123 456.';
    if (checkout.mode === 'envio') {
      if (CFG.zonas.length && !checkout.zona) e.zona = 'Elegí la zona de entrega.';
      if (checkout.addr.trim().length < 5) e.addr = 'Escribí la dirección de entrega.';
    }
    if (CFG.pagos.length && !checkout.pago) e.pago = 'Elegí cómo vas a pagar.';
    return e;
  }
  function viewCart() {
    setMeta('Tu pedido | La Plaza Perfumería'); setJsonLd(null);
    if (sent) {
      return '<div class="wrap page"><div class="done"><div class="done-icon">' + I.check + '</div><h1>Tu pedido está listo</h1>' +
        '<p>Tocá el botón para abrir WhatsApp con el pedido ya escrito. Lo enviás desde ahí y te confirmamos disponibilidad y total.</p>' +
        '<div class="row"><a class="btn" id="wa-send" href="' + waLink(orderMessage()) + '" target="_blank" rel="noopener">' + I.wa + 'Enviar por WhatsApp</a></div>' +
        '<div class="row"><button type="button" class="btn btn-ghost btn-sm" data-action="sent-ok">Ya lo envié, vaciar carrito</button><button type="button" class="link-btn" data-action="sent-back">Volver al pedido</button></div></div></div>';
    }
    if (!cart.length) return '<div class="wrap page"><div class="done"><h1>Tu carrito está vacío</h1><p>Elegí una fragancia o probá varias en decant.</p><div class="row"><a class="btn" href="#/catalogo">Descubrir perfumes</a><a class="btn btn-lilac" href="#/test">Encontrar mi fragancia</a></div></div></div>';
    var e = tried ? errors() : {};
    return '<div class="wrap page"><h1 class="page-title" style="margin-bottom:22px">Tu pedido</h1><div class="cart-layout">' +
      '<div class="cart-main"><div id="lines" style="display:flex;flex-direction:column;gap:10px"></div>' +
      '<form class="form" id="checkout" novalidate>' +
        '<h2>Tus datos</h2>' +
        field('name', 'Nombre y apellido', '<input type="text" id="c-name" name="name" autocomplete="name" value="' + esc(checkout.name) + '">', e) +
        field('phone', 'Teléfono / WhatsApp', '<input type="tel" id="c-phone" name="phone" autocomplete="tel" inputmode="tel" placeholder="0981 123 456" value="' + esc(checkout.phone) + '">', e) +
        '<fieldset class="seg"><legend>¿Cómo lo recibís?</legend>' +
          '<label><input type="radio" name="mode" value="envio"' + (checkout.mode === 'envio' ? ' checked' : '') + '>Envío</label>' +
          '<label><input type="radio" name="mode" value="retiro"' + (checkout.mode === 'retiro' ? ' checked' : '') + '>Retiro sin costo</label></fieldset>' +
        '<div id="delivery" class="form"></div>' +
        (CFG.pagos.length ? '<fieldset class="seg' + (e.pago ? ' has-error' : '') + '" id="pago-set"><legend>Pago</legend>' + CFG.pagos.map(function (m) {
          return '<label><input type="radio" name="pago" value="' + esc(m) + '"' + (checkout.pago === m ? ' checked' : '') + '>' + esc(m) + '</label>';
        }).join('') + '</fieldset><div id="err-pago">' + (e.pago ? '<span class="err" role="alert">' + e.pago + '</span>' : '') + '</div>' : '') +
        (DATA.cupones.length ? '<label>Cupón de descuento<span class="coupon"><input type="text" id="c-cupon" name="cuponInput" autocomplete="off" value="' + esc(checkout.cupon) + '" aria-describedby="cupon-msg"><button type="button" class="btn btn-ghost btn-sm" data-action="apply-coupon">Aplicar</button></span><span id="cupon-msg" class="fine"></span></label>' : '') +
        field('notes', 'Comentarios (opcional)', '<textarea id="c-notes" name="notes">' + esc(checkout.notes) + '</textarea>', e) +
      '</form></div>' +
      '<aside class="cart-side" aria-label="Resumen"><div class="summary" id="summary"></div>' +
        '<button type="submit" form="checkout" class="btn btn-block">' + I.wa + 'Continuar al pedido por WhatsApp</button>' +
        '<p class="fine">En el paso siguiente se abre WhatsApp con el pedido escrito y lo enviás vos. Tus datos no se guardan en este sitio. El pago se coordina por WhatsApp: no hay cobro online.</p>' +
      '</aside></div></div>';
  }
  function field(key, label, input, e) {
    var err = e && e[key];
    return '<label class="' + (err ? 'has-error' : '') + '" id="w-' + key + '">' + label + input + (err ? '<span class="err" role="alert">' + err + '</span>' : '') + '</label>';
  }
  function renderCartParts() {
    document.getElementById('lines').innerHTML = cart.map(function (i) { return lineHTML(i, 'c'); }).join('');
    var e = tried ? errors() : {};
    var del = document.getElementById('delivery');
    if (checkout.mode === 'envio') {
      del.innerHTML = (CFG.zonas.length ? '<label class="' + (e.zona ? 'has-error' : '') + '">Zona de entrega<select id="c-zona" name="zona"><option value="">Elegí tu zona</option>' + CFG.zonas.map(function (z) {
          return '<option value="' + esc(z.zona) + '"' + (checkout.zona === z.zona ? ' selected' : '') + '>' + esc(z.zona) + ': ' + (z.consultar ? 'costo a confirmar' : z.costo === 0 ? 'sin costo' : gs(z.costo)) + (z.tiempo ? ', ' + esc(z.tiempo).toLowerCase() : '') + '</option>';
        }).join('') + '</select>' + (e.zona ? '<span class="err" role="alert">' + e.zona + '</span>' : '') + '</label>' : '') +
        '<label class="' + (e.addr ? 'has-error' : '') + '">Dirección y referencia<input type="text" id="c-addr" name="addr" autocomplete="street-address" placeholder="Calle, número, barrio, ciudad" value="' + esc(checkout.addr) + '">' + (e.addr ? '<span class="err" role="alert">' + e.addr + '</span>' : '') + '</label>';
    } else del.innerHTML = '<p class="fine" style="font-size:14px">Retiro sin costo en ' + esc(CFG.retiro) + '. Coordinamos día y horario por WhatsApp.</p>';
    var pe = document.getElementById('err-pago'); if (pe) pe.innerHTML = e.pago ? '<span class="err" role="alert">' + e.pago + '</span>' : '';
    var ps = document.getElementById('pago-set'); if (ps) ps.classList.toggle('has-error', !!e.pago);
    renderSummary();
  }
  function renderSummary() {
    var t = totals();
    var cm = document.getElementById('cupon-msg');
    if (cm) cm.textContent = couponMsg || (t.cp ? (t.cp.msg || 'Cupón ' + t.cp.c.codigo + ' aplicado.') : '');
    document.getElementById('summary').innerHTML = '<h2>Resumen</h2>' +
      '<div class="sum-row"><span>Productos (' + units() + ')</span><span>' + gs(t.sub) + '</span></div>' +
      (t.saved ? '<div class="sum-row disc"><span>Ahorrás en ofertas</span><span>' + gs(t.saved) + '</span></div>' : '') +
      (t.disc ? '<div class="sum-row disc"><span>Cupón ' + esc(t.cp.c.codigo) + '</span><span>−' + gs(t.disc) + '</span></div>' : '') +
      '<div class="sum-row"><span>' + esc(t.sh.label) + '</span><span>' + (checkout.mode === 'retiro' ? 'Sin costo' : t.sh.known ? (t.sh.cost ? gs(t.sh.cost) : 'Sin costo') : 'A confirmar') + '</span></div>' +
      '<div class="sum-row sum-total"><span>Total</span><span>' + gs(t.total) + (t.sh.known ? '' : ' + envío') + '</span></div>';
  }
  function orderMessage() {
    var t = totals();
    var m = ['Hola La Plaza! Quiero hacer este pedido:', ''];
    cart.forEach(function (i) { var p = BY_ID[i.id], v = variant(p, i.v); m.push('• ' + p.marca + ' ' + p.nombre + ', ' + v.label + ' x' + i.qty + ': ' + gs(v.price * i.qty)); });
    m.push('', 'Subtotal: ' + gs(t.sub));
    if (t.disc) m.push('Cupón ' + t.cp.c.codigo + ': −' + gs(t.disc));
    if (checkout.mode === 'retiro') m.push('Entrega: Retiro en ' + CFG.retiro);
    else { m.push('Entrega: ' + t.sh.label + (t.sh.known ? ': ' + (t.sh.cost ? gs(t.sh.cost) : 'sin costo') : '')); m.push('Dirección: ' + checkout.addr.trim()); }
    m.push('Total: ' + gs(t.total) + (t.sh.known ? '' : ' + envío'));
    if (checkout.pago) m.push('Pago: ' + checkout.pago);
    m.push('', 'Nombre: ' + checkout.name.trim(), 'Teléfono: ' + checkout.phone.trim());
    if (checkout.notes.trim()) m.push('Comentarios: ' + checkout.notes.trim());
    return m.join('\n');
  }

  /* ---------- test de personalidad olfativa ---------- */
  var PERSONAS = [
    ['dulce', 'Dulce y cariñosa/o', 'Te gusta lo cálido y cercano', ['Dulce']],
    ['fresca', 'Fresca/o y relajada/o', 'Sencillez, aire libre, buena onda', ['Fresco']],
    ['elegante', 'Elegante y misteriosa/o', 'Detalles, noche, presencia', ['Oriental', 'Amaderado']],
    ['intensa', 'Intensa/o y con carácter', 'Dejás huella donde vas', ['Especiado', 'Amaderado']],
    ['romantica', 'Romántica/o', 'Flores, ternura, detalles', ['Floral']],
    ['alegre', 'Alegre y espontánea/o', 'Colores, fiestas, energía', ['Frutal', 'Dulce']]
  ];
  var NOTE_GROUPS = [
    ['vainilla', 'Vainilla y caramelo', ['vainilla', 'caramelo', 'praline', 'haba tonka', 'miel', 'azucar', 'leche', 'goloso']],
    ['citricos', 'Cítricos', ['limon', 'bergamota', 'mandarina', 'naranja', 'toronja', 'pomelo', 'citric']],
    ['maderas', 'Maderas', ['cedro', 'sandalo', 'madera', 'vetiver', 'oud', 'abeto', 'abedul', 'ciprés', 'cipres', 'pachuli', 'amaderad']],
    ['especias', 'Especias', ['canela', 'cardamomo', 'jengibre', 'azafran', 'especia', 'nuez moscada', 'pimienta']],
    ['flores', 'Flores', ['jazmin', 'rosa', 'azahar', 'iris', 'lirio', 'geranio', 'flor', 'orquidea', 'neroli', 'lavanda']],
    ['frutas', 'Frutas', ['manzana', 'pina', 'pera', 'fresa', 'frambuesa', 'durazno', 'coco', 'higo', 'ciruela', 'fruta', 'grosella', 'mburucuya', 'datiles']]
  ];
  var quiz = { step: 0, a: {} };
  function notesHay(p) { return norm(p.notas.salida + ' ' + p.notas.corazon + ' ' + p.notas.fondo); }
  function quizQuestions() {
    var P = DATA.productos;
    var fams = uniq(P.map(function (p) { return p.familia; }));
    var qs = [
      { key: 'perso', q: '¿Cómo te describirías?', opts: PERSONAS.filter(function (x) { return x[3].some(function (f) { return fams.indexOf(f) !== -1; }); }).map(function (x) { return [x[0], x[1], x[2]]; }) },
      { key: 'notas', q: '¿Qué notas te gustan más?', opts: NOTE_GROUPS.filter(function (g) { return P.some(function (p) { var h = notesHay(p); return g[2].some(function (k) { return h.indexOf(k) !== -1; }); }); }).map(function (g) { return [g[0], g[1], '']; }) }
    ];
    var oc = sortEs(uniq([].concat.apply([], P.map(function (p) { return p.ocasion; }))));
    if (oc.length) qs.push({ key: 'ocasion', q: '¿Para qué lo querés usar?', opts: oc.map(function (o) { return [o, o, '']; }) });
    if (P.some(function (p) { return p.clima; })) qs.push({ key: 'clima', q: '¿Para qué clima?', opts: [['Calor', 'Calor', 'Verano, días soleados'], ['Frío', 'Frío', 'Otoño, invierno, días lluviosos']] });
    qs.push({ key: 'genero', q: '¿Para quién es?', opts: [['Masculino', 'Para él', 'Incluye unisex'], ['Femenino', 'Para ella', 'Incluye unisex'], ['', 'Sin preferencia', '']] });
    qs.push({ key: 'budget', q: '¿Cuál es tu presupuesto?', opts: [['decant', 'Quiero probar en decant', 'Hasta Gs. 50.000'], ['200000', 'Hasta Gs. 200.000', ''], ['300000', 'Hasta Gs. 300.000', ''], ['', 'Sin límite', '']] });
    return qs;
  }
  function quizPrice(p, budget) {
    var vs = variants(p).filter(function (v) { return v.state !== 'out'; });
    if (budget === 'decant') vs = vs.filter(function (v) { return v.key !== 'botella' && v.price <= 50000; });
    else if (budget) vs = vs.filter(function (v) { return v.price <= Number(budget); });
    return vs.length ? Math.min.apply(null, vs.map(function (v) { return v.price; })) : null;
  }
  function quizResults() {
    var a = quiz.a;
    var persona = PERSONAS.filter(function (x) { return x[0] === a.perso; })[0];
    var group = NOTE_GROUPS.filter(function (g) { return g[0] === a.notas; })[0];
    var scored = DATA.productos.filter(function (p) {
      if (a.genero && p.genero !== a.genero && p.genero !== 'Unisex') return false;
      return quizPrice(p, a.budget) != null;
    }).map(function (p) {
      var why = [], score = 0, need = 0, got = 0;
      if (persona) { need++; if (persona[3].indexOf(p.familia) !== -1) { score += 3; got++; why.push('familia ' + p.familia.toLowerCase()); } }
      if (group) { need++; var h = notesHay(p); if (group[2].some(function (k) { return h.indexOf(k) !== -1; })) { score += 2; got++; why.push('tiene notas de ' + group[1].toLowerCase()); } }
      if (a.ocasion) { need++; if (p.ocasion.indexOf(a.ocasion) !== -1) { score += 2; got++; why.push('ideal para ' + a.ocasion.toLowerCase()); } }
      if (a.clima) { need++; if (p.clima === a.clima || p.clima === 'Todo el año') { score += 1; got++; why.push(p.clima === 'Todo el año' ? 'va todo el año' : 'pensado para clima ' + a.clima.toLowerCase()); } }
      var price = quizPrice(p, a.budget);
      if (a.budget === 'decant') why.push('decant desde ' + gs(price));
      else if (a.budget) why.push('dentro de tu presupuesto');
      return { p: p, score: score, full: got === need, why: why };
    }).sort(function (x, y) { return (y.score - x.score) || (y.p.destacado - x.p.destacado) || fromPrice(x.p) - fromPrice(y.p); });
    var exact = scored.filter(function (x) { return x.full; });
    var close = !exact.length;
    var list = (exact.length ? exact : scored.filter(function (x) { return x.score > 0; })).slice(0, 4);
    if (!list.length) list = scored.slice(0, 4);
    if (!list.length) list = DATA.productos.filter(function (p) { return p.destacado; }).slice(0, 4).map(function (p) { return { p: p, why: [] }; });
    return { close: close, list: list, persona: persona };
  }
  function viewQuiz() {
    setMeta('Test de personalidad olfativa | La Plaza Perfumería', 'Respondé unas preguntas y te recomendamos perfumes de nuestro catálogo según tu estilo y presupuesto.'); setJsonLd(null);
    var qs = quizQuestions();
    if (quiz.step < qs.length) {
      var q = qs[quiz.step];
      return '<div class="wrap page"><div class="quiz">' +
        '<div><span class="eyebrow">Test de personalidad olfativa</span></div>' +
        '<div class="quiz-card"><div class="progress" role="progressbar" aria-label="Progreso del test" aria-valuemin="1" aria-valuemax="' + qs.length + '" aria-valuenow="' + (quiz.step + 1) + '">' + qs.map(function (_, i) { return '<span class="' + (i <= quiz.step ? 'on' : '') + '"></span>'; }).join('') + '</div>' +
        '<span class="muted" style="font-size:13px;font-weight:600">Pregunta ' + (quiz.step + 1) + ' de ' + qs.length + '</span>' +
        '<h1 tabindex="-1" id="quiz-h">' + q.q + '</h1>' +
        '<div class="opts">' + q.opts.map(function (o) { return '<button type="button" class="opt" data-action="answer" data-v="' + esc(o[0]) + '"><b>' + esc(o[1]) + '</b>' + (o[2] ? '<span>' + esc(o[2]) + '</span>' : '') + '</button>'; }).join('') + '</div>' +
        '<div class="quiz-nav">' + (quiz.step > 0 ? '<button type="button" class="btn btn-ghost btn-sm" data-action="q-back">Atrás</button>' : '') + '<button type="button" class="link-btn" data-action="q-skip">Saltar pregunta</button></div></div>' +
      '</div></div>';
    }
    var r = quizResults();
    var fam = r.persona ? r.persona[3].filter(function (f) { return FAMILIES[f]; })[0] : null;
    return '<div class="wrap page"><div class="quiz" style="max-width:var(--wrap)">' +
      '<div class="profile"><span class="eyebrow">Tu resultado</span><h1 tabindex="-1" id="quiz-h">' + (r.persona ? 'Tu estilo: ' + r.persona[3].join(' / ').toLowerCase() : (r.close ? 'Opciones cercanas' : 'Te pueden gustar')) + '</h1>' +
        '<p>' + (fam ? FAMILIES[fam].d + ' ' : '') + (r.close ? 'No encontramos una coincidencia exacta con todas tus respuestas; estas son las más cercanas en nuestro catálogo.' : 'Estas fragancias de nuestro catálogo coinciden con tus respuestas. Si dudás, empezá por un decant.') + '</p></div>' +
      '<div class="grid">' + r.list.map(function (x) { return card(x.p, a_mode(), x.why.length ? '<span class="why">Por qué: ' + esc(x.why.join(', ')) + '</span>' : ''); }).join('') + '</div>' +
      '<div class="quiz-nav"><button type="button" class="btn btn-ghost" data-action="q-restart">Repetir el test</button><a class="btn" href="#/catalogo">Descubrir perfumes</a></div>' +
    '</div></div>';
  }
  function a_mode() { return quiz.a.budget === 'decant' ? 'decants' : ''; }

  /* ---------- comparador ---------- */
  function viewCompare() {
    setMeta('Comparar perfumes | La Plaza Perfumería'); setJsonLd(null);
    var list = compare.map(function (id) { return BY_ID[id]; });
    if (list.length < 2) {
      return '<div class="wrap page"><h1 class="page-title">Comparar perfumes</h1><div class="empty"><h2>' + (list.length ? 'Sumá otro perfume' : 'Elegí perfumes para comparar') + '</h2><p>Tocá el ícono de comparar en hasta 3 perfumes y vas a ver sus diferencias lado a lado.</p><a class="btn btn-sm" href="#/catalogo">Descubrir perfumes</a></div></div>';
    }
    function rowIf(label, fn) {
      var vals = list.map(fn);
      if (!vals.some(function (v) { return v; })) return '';
      return '<tr><th scope="row">' + label + '</th>' + vals.map(function (v) { return '<td>' + (v || '-') + '</td>'; }).join('') + '</tr>';
    }
    return '<div class="wrap page"><h1 class="page-title" style="margin-bottom:20px">Comparar perfumes</h1><div class="table-wrap"><table class="cmp-table">' +
      '<tr><th scope="row">Perfume</th>' + list.map(function (p) {
        return '<td><div class="cmp-prod"><a class="line-media" href="#/p/' + p.id + '">' + img(p) + '</a><span class="line-brand">' + esc(p.marca) + '</span><a class="line-name" href="#/p/' + p.id + '">' + esc(p.nombre) + '</a>' +
          '<button type="button" class="link-btn" style="align-self:flex-start;padding-left:0" data-action="cmp" data-id="' + p.id + '" aria-pressed="true">Quitar</button></div></td>';
      }).join('') + '</tr>' +
      rowIf('Precio', function (p) { var pi = priceInfo(p); return '<b>' + (pi.many ? 'desde ' : '') + gs(pi.price) + '</b>' + (pi.old ? ' <s class="muted">' + gs(pi.old) + '</s>' : ''); }) +
      rowIf('Presentaciones', function (p) { return esc(presLabel(p)); }) +
      rowIf('Disponibilidad', function (p) { var a = avail(p); return '<span class="avail ' + a.cls + '">' + a.txt + '</span>'; }) +
      rowIf('Género', function (p) { return esc(p.genero); }) +
      rowIf('Familia', function (p) { return esc(p.familia); }) +
      rowIf('Notas de salida', function (p) { return esc(p.notas.salida); }) +
      rowIf('Notas de corazón', function (p) { return esc(p.notas.corazon); }) +
      rowIf('Notas de fondo', function (p) { return esc(p.notas.fondo); }) +
      rowIf('Momento', function (p) { return esc(p.momento); }) +
      rowIf('Clima', function (p) { return esc(p.clima); }) +
      rowIf('Ocasión', function (p) { return esc(p.ocasion.join(', ')); }) +
      rowIf('Intensidad', function (p) { return esc(p.intensidad); }) +
      rowIf('Duración', function (p) { return esc(p.duracion); }) +
      rowIf('Concentración', function (p) { return esc(p.concentracion); }) +
      '<tr><th scope="row"></th>' + list.map(function (p) { return '<td>' + (avail(p).cls === 'ok' ? '<button type="button" class="btn btn-sm" data-action="quick" data-id="' + p.id + '">Agregar</button>' : '<a class="btn btn-sm btn-ghost" href="#/p/' + p.id + '">Ver</a>') + '</td>'; }).join('') + '</tr>' +
      '</table></div></div>';
  }

  /* ---------- ayuda ---------- */
  function viewHelp() {
    setMeta('Compras y envíos | La Plaza Perfumería'); setJsonLd(null);
    return '<div class="wrap page help" style="max-width:820px"><h1 class="page-title">Compras y envíos</h1>' +
      '<section id="envios"><h2>Envíos y retiro</h2><p>Retiro sin costo en ' + esc(CFG.retiro) + ', coordinando día y horario por WhatsApp.</p>' +
        (CFG.zonas.length ? '<div class="summary" style="margin-top:14px">' + CFG.zonas.map(function (z) { return '<div class="sum-row"><span>' + esc(z.zona) + (z.tiempo ? ' · ' + esc(z.tiempo) : '') + '</span><span>' + (z.consultar ? 'Consultar' : z.costo === 0 ? 'Sin costo' : gs(z.costo)) + '</span></div>'; }).join('') + '</div>' : '') + '</section>' +
      '<section id="pagos"><h2>Pagos</h2><p>Aceptamos ' + esc(CFG.pagos.join(' y ').toLowerCase()) + '. El pago se coordina por WhatsApp al confirmar el pedido; este sitio no cobra en línea.</p></section>' +
      '<section id="decants"><h2>¿Qué es un decant?</h2><p>Es una porción del perfume original, trasvasada de la botella a un atomizador más chico. No es una botella sellada: es el mismo perfume en menor cantidad, para probarlo antes de comprar la botella.</p></section>' +
      '<section id="cambios"><h2>Cambios y devoluciones</h2><p>Escribinos por WhatsApp antes de confirmar tu compra y te contamos las condiciones para tu pedido.</p></section>' +
      '<section id="terminos"><h2>Términos de compra</h2><p>Los precios están en guaraníes. Al recibir tu pedido por WhatsApp te confirmamos disponibilidad, costo de envío y total; el pedido queda confirmado cuando acordamos el pago. Los cupones de descuento se aplican sobre el subtotal de productos y no son acumulables salvo que se indique.</p></section>' +
      '<section id="privacidad"><h2>Privacidad</h2><p>Este sitio no guarda tus datos personales en ningún servidor. Tu carrito, favoritos y comparador quedan solo en tu navegador. Los datos del pedido viajan únicamente en el mensaje de WhatsApp que vos decidís enviar.</p></section>' +
      '<div style="margin-top:28px"><a class="btn" href="' + waLink('Hola! Tengo una consulta.') + '" target="_blank" rel="noopener">' + I.wa + 'Escribinos por WhatsApp</a></div></div>';
  }
  function viewNotFound() {
    setMeta('No encontrado | La Plaza Perfumería');
    return '<div class="wrap page"><div class="done"><h1>No encontramos esa página</h1><p>Puede que el perfume ya no esté en el catálogo.</p><div class="row"><a class="btn" href="#/catalogo">Descubrir perfumes</a></div></div></div>';
  }

  /* ---------- router ---------- */
  function route() {
    closeLayer(true);
    clearInterval(heroTimer); heroTimer = null;
    var h = location.hash.replace(/^#/, '') || '/';
    var qi = h.indexOf('?');
    var path = qi === -1 ? h : h.slice(0, qi), qs = qi === -1 ? '' : h.slice(qi + 1);
    var html, after = null;
    if (path === '/' || path === '') { html = viewHome(); after = afterHome; }
    else if (path === '/catalogo') { html = viewCatalog(qs, ''); after = function () { renderCatalogParts(); }; }
    else if (path === '/decants') { html = viewCatalog(qs, 'decants'); after = function () { renderCatalogParts(); }; }
    else if (path.indexOf('/p/') === 0) { html = viewProduct(decodeURIComponent(path.slice(3))); if (pdp) after = function () { renderGallery(); renderPdpParts(); }; }
    else if (path === '/carrito') { html = viewCart(); if (cart.length && !sent) after = renderCartParts; }
    else if (path === '/test') html = viewQuiz();
    else if (path === '/comparar') html = viewCompare();
    else if (path === '/ayuda') html = viewHelp();
    else html = viewNotFound();
    app.innerHTML = '<div class="view">' + html + '</div>';
    if (after) after();
    // Entrada escalonada solo al llegar a una página (no al filtrar ni ordenar).
    app.querySelectorAll('.grid').forEach(function (g) {
      [].forEach.call(g.children, function (c, i) { c.style.setProperty('--i', i); c.classList.add('enter'); });
    });
    document.querySelectorAll('.nav a').forEach(function (a) {
      if (a.getAttribute('href').replace(/^#/, '') === h) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    document.getElementById('nav').classList.remove('open');
    document.getElementById('menu-btn').setAttribute('aria-expanded', 'false');
    document.getElementById('wa-float').hidden = path === '/carrito';
    renderCompareBar();
    var s = new URLSearchParams(qs).get('s');
    if (s && document.getElementById(s)) document.getElementById(s).scrollIntoView(); else window.scrollTo(0, 0);
  }

  /* ---------- eventos ---------- */
  document.addEventListener('click', function (ev) {
    var el = ev.target.closest('[data-action]');
    if (!el) {
      if (ev.target.id === 'overlay') closeLayer();
      return;
    }
    var a = el.getAttribute('data-action'), id = el.getAttribute('data-id'), v = el.getAttribute('data-v');
    if (a === 'hero') { setHero(Number(el.getAttribute('data-k'))); startHero(); return; }
    if (a === 'ml') { setMeter(Number(el.getAttribute('data-ml'))); return; }
    if (a === 'close') closeLayer();
    else if (a === 'close-nav') closeLayer(true);
    else if (a === 'fav') {
      var on = toggleFav(id);
      document.querySelectorAll('[data-action="fav"][data-id="' + id + '"]').forEach(function (b) { b.setAttribute('aria-pressed', on); b.setAttribute('aria-label', (on ? 'Quitar de' : 'Agregar a') + ' favoritos'); });
      toast(on ? 'Guardado en favoritos' : 'Quitado de favoritos');
      if (cat && cat.fav && document.getElementById('results')) renderCatalogParts();
    } else if (a === 'cmp') {
      var was = compare.indexOf(id) !== -1;
      var now = toggleCompare(id);
      if (now !== was) toast(now ? 'Agregado al comparador' : 'Quitado del comparador');
      if (location.hash.indexOf('#/comparar') === 0) route();
    } else if (a === 'cmp-clear') { compare = []; save('lp_compare', compare); renderCompareBar(); }
    else if (a === 'quick') openQuick(id, el.getAttribute('data-mode') || '');
    else if (a === 'qk-inc' || a === 'qk-dec') { quick.qty = Math.max(1, quick.qty + (a === 'qk-inc' ? 1 : -1)); refreshSheet('[data-action="' + a + '"]'); }
    else if (a === 'qk-add') {
      var qp = BY_ID[quick.id], qv = variant(qp, quick.v), n1 = addToCart(quick.id, quick.v, quick.qty);
      if (n1) openDrawer('Agregaste ' + qp.nombre + ' · ' + qv.label + (n1 > 1 ? ' × ' + n1 : ''));
      else toast('No hay más stock de esta presentación.');
    } else if (a === 'open-cart') openDrawer();
    else if (a === 'unchip') {
      var key = el.getAttribute('data-key');
      if (key === 'q') { cat.q = ''; var qi = document.getElementById('cat-q'); if (qi) qi.value = ''; }
      else if (key === 'fav') cat.fav = false;
      else cat[key] = cat[key].filter(function (x) { return x !== v; });
      cat.limit = PAGE_SIZE; syncCatalog();
    } else if (a === 'clear') {
      var keep = cat.sort, mode = cat.mode; cat = parseQuery('', mode); cat.sort = keep;
      var q2 = document.getElementById('cat-q'); if (q2) q2.value = '';
      syncCatalog();
    } else if (a === 'more') {
      var before = cat.limit; cat.limit += PAGE_SIZE; renderCatalogParts();
      var next = document.querySelectorAll('#results .card')[before]; if (next) { var lnk = next.querySelector('a.card-name'); if (lnk) lnk.focus(); }
    } else if (a === 'zoom') { var zp = BY_ID[pdp.id]; openLightbox('img/' + zp.fotos[pdp.foto], zp.marca + ' ' + zp.nombre); }
    else if (a === 'thumb') { pdp.foto = Number(el.getAttribute('data-i')); renderGallery(); }
    else if (a === 'pdp-inc' || a === 'pdp-dec') {
      pdp.qty = Math.max(1, pdp.qty + (a === 'pdp-inc' ? 1 : -1)); renderPdpParts();
      var b2 = document.querySelector('#buy-row [data-action="' + a + '"]'); if (b2 && !b2.disabled) b2.focus();
    } else if (a === 'add') {
      var p = BY_ID[pdp.id], cur = variant(p, pdp.v), n2 = addToCart(p.id, pdp.v, pdp.qty);
      if (n2) { pdp.qty = 1; renderPdpParts(); openDrawer('Agregaste ' + p.nombre + ' · ' + cur.label + (n2 > 1 ? ' × ' + n2 : '')); }
      else toast('No hay más stock de esta presentación.');
    } else if (a === 'copy') {
      var url = shareUrl(BY_ID[pdp.id]);
      var done = function () { toast('Enlace copiado'); };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(url).then(done, function () { prompt0(url); });
      else prompt0(url);
    } else if (a === 'native-share') {
      var sp = BY_ID[pdp.id];
      navigator.share({ title: sp.marca + ' ' + sp.nombre, text: 'Mirá este perfume en La Plaza', url: shareUrl(sp) }).catch(function () {});
    } else if (a === 'rm') {
      removeLine(id, v);
      if (layer === 'drawer') refreshDrawer();
      if (location.hash.indexOf('#/carrito') === 0) { if (!cart.length) route(); else renderCartParts(); }
    } else if (a === 'q-inc' || a === 'q-dec') {
      var line = lineOf(id, v);
      if (line) setQty(id, v, line.qty + (a === 'q-inc' ? 1 : -1));
      if (layer === 'drawer') refreshDrawer(); else if (location.hash.indexOf('#/carrito') === 0) { renderCartParts(); var b3 = document.querySelector('#lines [data-action="' + a + '"][data-id="' + id + '"][data-v="' + v + '"]'); if (b3 && !b3.disabled) b3.focus(); }
    } else if (a === 'apply-coupon') {
      var code = (document.getElementById('c-cupon').value || '').trim().toUpperCase();
      var found = DATA.cupones.filter(function (c) { return c.codigo === code; })[0];
      checkout.cupon = found ? code : '';
      couponMsg = !code ? '' : found ? '' : 'Ese código no es válido o ya no está vigente.';
      save('lp_checkout', checkout); renderSummary();
    } else if (a === 'sent-ok') { cart = []; save('lp_cart', cart); updateBadges(); sent = false; tried = false; checkout.cupon = ''; save('lp_checkout', checkout); route(); }
    else if (a === 'sent-back') { sent = false; route(); }
    else if (a === 'answer' || a === 'q-skip') {
      var k = quizQuestions()[quiz.step].key;
      if (a === 'answer' && v) quiz.a[k] = v; else delete quiz.a[k];
      quiz.step++; route(); focusQuiz();
    } else if (a === 'q-back') { quiz.step = Math.max(0, quiz.step - 1); route(); focusQuiz(); }
    else if (a === 'q-restart') { quiz = { step: 0, a: {} }; route(); focusQuiz(); }
  });
  function prompt0(url) {
    var i = document.createElement('input'); i.value = url; document.body.appendChild(i); i.select();
    try { document.execCommand('copy'); toast('Enlace copiado'); } catch (e) { toast(url); }
    i.remove();
  }
  function focusQuiz() { var h = document.getElementById('quiz-h'); if (h) h.focus(); }

  document.addEventListener('pointerover', function (ev) {
    var r = ev.target.closest && ev.target.closest('.fi-row[data-fam]');
    if (r && window.matchMedia('(hover: hover)').matches) setFam(r.getAttribute('data-fam'));
  });
  document.addEventListener('focusin', function (ev) {
    var r = ev.target.closest && ev.target.closest('.fi-row[data-fam]');
    if (r) setFam(r.getAttribute('data-fam'));
  });
  document.addEventListener('change', function (ev) {
    var t = ev.target;
    if (t.matches('[data-filter]')) {
      var k = t.getAttribute('data-filter');
      if (t.checked) cat[k].push(t.value); else cat[k] = cat[k].filter(function (x) { return x !== t.value; });
      cat.limit = PAGE_SIZE; syncCatalog(t.id);
    } else if (t.id === 'cat-sort') { cat.sort = t.value; syncCatalog(); }
    else if (t.name === 'variant') { pdp.v = t.value; pdp.qty = 1; renderPdpParts(t.value); }
    else if (t.name === 'qv') { quick.v = t.value; quick.qty = 1; refreshSheet('input[value="' + t.value + '"]'); }
    else if (t.matches('[data-action="swap"]')) {
      var sid = t.getAttribute('data-id');
      changeVariant(sid, t.getAttribute('data-v'), t.value);
      var nid = 'sw-' + (layer === 'drawer' ? 'd' : 'c') + '-' + sid + '-' + t.value;
      if (layer === 'drawer') { document.getElementById('drawer').innerHTML = drawerHTML(); }
      else if (location.hash.indexOf('#/carrito') === 0) renderCartParts();
      var ne = document.getElementById(nid); if (ne) ne.focus();
    } else if (t.form && t.form.id === 'checkout') {
      if (t.name === 'mode' || t.name === 'pago' || t.name === 'zona') {
        checkout[t.name] = t.value; save('lp_checkout', checkout);
        if (t.name === 'mode') { renderCartParts(); var r = document.querySelector('input[name="mode"]:checked'); if (r) r.focus(); }
        else if (t.name === 'zona') renderSummary();
        else if (tried) renderCartParts();
      }
    }
  });
  var searchTimer;
  document.addEventListener('input', function (ev) {
    var t = ev.target;
    if (t.id === 'cat-q') { clearTimeout(searchTimer); searchTimer = setTimeout(function () { cat.q = t.value.trim(); cat.limit = PAGE_SIZE; syncCatalog(); }, 200); }
    else if (t.form && t.form.id === 'checkout' && ['name', 'phone', 'addr', 'notes'].indexOf(t.name) !== -1) { checkout[t.name] = t.value; save('lp_checkout', checkout); }
  });
  document.addEventListener('submit', function (ev) {
    if (ev.target.id === 'search') {
      ev.preventDefault();
      var q = ev.target.q.value.trim();
      location.hash = '#/catalogo' + (q ? '?q=' + encodeURIComponent(q) : '');
      ev.target.q.value = '';
      ev.target.classList.remove('open'); document.getElementById('search-btn').setAttribute('aria-expanded', 'false');
    } else if (ev.target.id === 'checkout') {
      ev.preventDefault();
      tried = true;
      var keys = Object.keys(errors());
      if (keys.length) {
        route();
        var target = { name: 'c-name', phone: 'c-phone', zona: 'c-zona', addr: 'c-addr', pago: 'pago-set' }[keys[0]];
        var el = document.getElementById(target);
        if (el) { el.scrollIntoView({ block: 'center' }); var f = el.matches('input,select') ? el : el.querySelector('input'); if (f) f.focus(); }
        return;
      }
      sent = true; route();
      var s = document.getElementById('wa-send'); if (s) s.focus();
    }
  });
  document.addEventListener('keydown', function (ev) {
    if (!layer) return;
    if (ev.key === 'Escape') { closeLayer(); return; }
    if (ev.key === 'Tab') {
      var box = document.getElementById(layer);
      var f = [].slice.call(box.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), select, textarea')).filter(function (x) { return x.offsetParent !== null; });
      if (!f.length) return;
      if (ev.shiftKey && document.activeElement === f[0]) { ev.preventDefault(); f[f.length - 1].focus(); }
      else if (!ev.shiftKey && document.activeElement === f[f.length - 1]) { ev.preventDefault(); f[0].focus(); }
    }
  });
  document.getElementById('overlay').addEventListener('click', function () { closeLayer(); });
  document.getElementById('cart-btn').addEventListener('click', function () { openDrawer(); });
  document.getElementById('menu-btn').addEventListener('click', function () {
    var open = document.getElementById('nav').classList.toggle('open');
    this.setAttribute('aria-expanded', open);
  });
  document.getElementById('search-btn').addEventListener('click', function () {
    var f = document.getElementById('search');
    var open = f.classList.toggle('open');
    this.setAttribute('aria-expanded', open);
    if (open) document.getElementById('search-q').focus();
  });
  window.addEventListener('hashchange', function () { if (location.hash.indexOf('#/carrito') !== 0) sent = false; route(); });

  /* ---------- arranque ---------- */
  if (!location.hash && window.INITIAL_ROUTE) history.replaceState(null, '', '#' + window.INITIAL_ROUTE);
  fetch('data/catalogo.json', { cache: 'no-cache' })
    .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then(function (d) {
      DATA = d; CFG = d.config; DATA.cupones = d.cupones || [];
      d.productos.forEach(function (p) { BY_ID[p.id] = p; });
      if (checkout.cupon && !DATA.cupones.some(function (c) { return c.codigo === checkout.cupon; })) checkout.cupon = '';
      cleanState(); updateBadges(); renderChrome(); route();
    })
    .catch(function () {
      app.innerHTML = '<div class="wrap page"><div class="done"><h1>No pudimos cargar el catálogo</h1><p>Revisá tu conexión y volvé a intentar.</p><div class="row"><button class="btn" type="button" onclick="location.reload()">Reintentar</button></div></div></div>';
    });
})();
