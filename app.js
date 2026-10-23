(function () {
  "use strict";
  var KEY = "majd-data-v1", META = "majd-meta-v1";
  var rows = [], meta = {};
  var $ = function (id) { return document.getElementById(id); };

  /* ---------- تخزين ---------- */
  function load() {
    try { rows = JSON.parse(localStorage.getItem(KEY) || "[]"); } catch (e) { rows = []; }
    try { meta = JSON.parse(localStorage.getItem(META) || "{}"); } catch (e) { meta = {}; }
  }
  function save() {
    try {
      localStorage.setItem(KEY, JSON.stringify(rows));
      localStorage.setItem(META, JSON.stringify(meta));
      return true;
    } catch (e) { return false; }
  }

  /* ---------- مساعدات ---------- */
  function norm(s) { return String(s == null ? "" : s).trim().toUpperCase().replace(/\s+/g, ""); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function fmtPrice(p) {
    var n = parseFloat(String(p).replace(/,/g, ""));
    return isNaN(n) ? esc(p) : n.toLocaleString("en-US", { maximumFractionDigits: 2 });
  }
  function toast(t) {
    var m = $("msg"); m.textContent = t; m.hidden = false;
    clearTimeout(toast.t); toast.t = setTimeout(function () { m.hidden = true; }, 6000);
  }

  /* ---------- قراءة الإكسيل ---------- */
  // يدور على أعمدة المصنع / الموديل / السعر بالاسم، ولو ملقاهاش بياخد A و B و C
  function parseMatrix(m) {
    var fi = 0, mi = 1, pi = 2, start = 0, i, j;
    for (i = 0; i < Math.min(m.length, 10); i++) {
      var r = m[i] || [], f = -1, mo = -1, p = -1;
      for (j = 0; j < r.length; j++) {
        var h = String(r[j] == null ? "" : r[j]).trim().toLowerCase();
        if (/مصنع|factory|maker|brand/.test(h)) f = j;
        else if (/موديل|model|كود|code/.test(h)) mo = j;
        else if (/سعر|price|cost/.test(h)) p = j;
      }
      if (mo > -1) { fi = f > -1 ? f : 0; mi = mo; pi = p > -1 ? p : 2; start = i + 1; break; }
    }
    var out = [], seen = {};
    for (i = start; i < m.length; i++) {
      var a = m[i]; if (!a) continue;
      var model = String(a[mi] == null ? "" : a[mi]).trim();
      if (!model) continue;
      var item = { f: String(a[fi] == null ? "" : a[fi]).trim(), m: model, p: a[pi] == null ? "" : a[pi] };
      var key = norm(item.f) + "|" + norm(model);
      if (seen[key] != null) out[seen[key]] = item; else { seen[key] = out.length; out.push(item); }
    }
    return out;
  }
  function parseWorkbookBuffer(buf) {
    var wb = XLSX.read(new Uint8Array(buf), { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    return parseMatrix(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }));
  }
  function applyRows(n, source) {
    if (!n.length) { toast("ما لقيتش بيانات صالحة. تأكد إن الأعمدة: اسم المصنع، رقم الموديل، السعر."); return; }
    rows = n; meta = { at: Date.now(), source: source, count: n.length };
    var ok = save();
    fillFactories(); renderSearch(); renderStat(); invRefresh();
    toast(ok ? "تم تحديث البيانات: " + n.length + " موديل." : "اتحدثت البيانات لكن مقدرتش أحفظها على الموبايل.");
  }

  /* ---------- صفحة البحث ---------- */
  function factories() {
    var s = {}, out = [];
    rows.forEach(function (r) { var k = r.f || "بدون مصنع"; if (!s[k]) { s[k] = 1; out.push(k); } });
    return out.sort(function (a, b) { return a.localeCompare(b, "ar"); });
  }
  var facList = [];
  function setFactory(v) {
    $("factory").value = v;
    $("comboLabel").textContent = v || "كل المصانع (" + facList.length + ")";
    closeCombo(); renderSearch();
  }
  function renderCombo() {
    var q = norm($("comboQ").value);
    var items = [""].concat(facList).filter(function (f) { return !q || (f && norm(f).indexOf(q) > -1); });
    var cur = $("factory").value;
    $("comboList").innerHTML = items.length ? items.map(function (f) {
      return '<li role="option" data-v="' + esc(f) + '"' + (f === cur ? ' class="sel"' : "") + ">" + (f ? esc(f) : "كل المصانع") + "</li>";
    }).join("") : '<li class="none">لا يوجد مصنع بهذا الاسم</li>';
  }
  function openCombo() {
    $("comboPanel").hidden = false; $("comboBtn").setAttribute("aria-expanded", "true");
    $("comboQ").value = ""; renderCombo(); $("comboQ").focus();
  }
  function closeCombo() { $("comboPanel").hidden = true; $("comboBtn").setAttribute("aria-expanded", "false"); }
  function fillFactories() {
    facList = factories();
    if (facList.indexOf($("factory").value) < 0) $("factory").value = "";
    $("comboLabel").textContent = $("factory").value || "كل المصانع (" + facList.length + ")";
    renderCombo();
  }
  function hl(model, q) {
    var raw = String(model), j = raw.toUpperCase().indexOf(q);
    if (!q || j < 0) return esc(raw);
    return esc(raw.slice(0, j)) + "<mark>" + esc(raw.slice(j, j + q.length)) + "</mark>" + esc(raw.slice(j + q.length));
  }
  function hlText(t, q) {
    var raw = String(t), k = q ? raw.toUpperCase().indexOf(q) : -1;
    if (k < 0) return esc(raw);
    return esc(raw.slice(0, k)) + "<mark>" + esc(raw.slice(k, k + q.length)) + "</mark>" + esc(raw.slice(k + q.length));
  }
  function card(r, q) {
    return '<div class="hit"><div class="c-f">' + hlText(r.f || "-", q) + '</div>' +
      '<div class="c-m">' + hlText(r.m, q) + '</div>' +
      '<div class="c-p">' + fmtPrice(r.p) + "</div></div>";
  }
  var HEAD = '<div class="hit head"><div>المصنع</div><div>رقم الموديل</div><div>السعر</div></div>';
  function renderSearch() {
    var out = $("results"), info = $("info");
    if (!rows.length) {
      info.textContent = "";
      out.innerHTML = '<div class="empty">لا توجد بيانات. افتح صفحة الإعدادات وحدّث من ملف إكسيل.</div>';
      return;
    }
    var fac = $("factory").value, q = norm($("q").value);
    var pool = fac ? rows.filter(function (r) { return (r.f || "بدون مصنع") === fac; }) : rows;
    if (!q) {
      info.textContent = fac ? "عدد موديلات المصنع: " + pool.length : "اختار مصنع أو اكتب رقم الموديل أو اسم المصنع";
      if (!fac) { out.innerHTML = ""; return; }
      out.innerHTML = HEAD + pool.slice(0, 60).map(function (r) { return card(r, ""); }).join("") +
        (pool.length > 60 ? '<div class="info">أول 60 موديل. اكتب رقم الموديل لتضييق البحث.</div>' : "");
      return;
    }
    var exact = pool.filter(function (r) { return norm(r.m) === q; });
    var list = exact.length ? exact : pool.filter(function (r) { return norm(r.m).indexOf(q) > -1 || norm(r.f).indexOf(q) > -1; });
    if (!list.length) { info.textContent = ""; out.innerHTML = '<div class="empty bad">غير موجود</div>'; return; }
    info.textContent = exact.length ? "" : "نتائج (" + list.length + ")";
    out.innerHTML = HEAD + list.slice(0, 50).map(function (r) { return card(r, q); }).join("");
  }

  /* ---------- صفحة الإعدادات ---------- */
  function renderStat() {
    if (!rows.length) { $("stat").textContent = "لا توجد بيانات محفوظة."; return; }
    var d = meta.at ? new Date(meta.at).toLocaleString("ar-EG") : "";
    $("stat").textContent = rows.length + " موديل، " + factories().length + " مصنع. آخر تحديث: " + d + (meta.source ? " (" + meta.source + ")" : "");
  }
  $("file").addEventListener("change", function (e) {
    var f = e.target.files[0]; if (!f) return;
    var rd = new FileReader();
    rd.onload = function () {
      try { applyRows(parseWorkbookBuffer(rd.result), f.name); }
      catch (err) { toast("تعذر قراءة الملف. تأكد إنه xlsx أو csv سليم."); }
    };
    rd.readAsArrayBuffer(f); e.target.value = "";
  });
  $("fromRepo").addEventListener("click", function () {
    toast("جارٍ التحميل…");
    fetch("data.xlsx?t=" + Date.now(), { cache: "no-store" })
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.arrayBuffer(); })
      .then(function (b) { applyRows(parseWorkbookBuffer(b), "data.xlsx"); })
      .catch(function () { toast("ما لقيتش data.xlsx أو مفيش اتصال بالإنترنت."); });
  });
  var clearArmed = false;
  $("clear").addEventListener("click", function () {
    if (!clearArmed) { clearArmed = true; $("clearAsk").hidden = false; setTimeout(function () { clearArmed = false; $("clearAsk").hidden = true; }, 5000); return; }
    rows = []; meta = {}; save(); fillFactories(); renderSearch(); renderStat(); invRefresh();
    clearArmed = false; $("clearAsk").hidden = true; toast("تم مسح البيانات.");
  });
  $("checkApp").addEventListener("click", function () {
    if (!("serviceWorker" in navigator)) { toast("المتصفح ده مش بيدعم التحديث التلقائي."); return; }
    navigator.serviceWorker.getRegistration().then(function (reg) {
      if (!reg) { toast("التطبيق شغال أونلاين بدون تخزين. أعد فتحه من الإنترنت."); return; }
      return reg.update().then(function () { toast("تم الفحص. لو فيه نسخة جديدة هتتفعل عند إعادة فتح التطبيق."); });
    }).catch(function () { toast("تعذر الفحص. تأكد من الاتصال بالإنترنت."); });
  });

  /* ---------- التبويبات ---------- */
  document.querySelectorAll(".tabs button").forEach(function (b) {
    b.addEventListener("click", function () {
      document.querySelectorAll(".tabs button").forEach(function (x) { x.classList.toggle("on", x === b); });
      $("view-search").hidden = b.dataset.view !== "search";
      $("view-invoice").hidden = b.dataset.view !== "invoice";
      $("view-settings").hidden = b.dataset.view !== "settings";
      window.scrollTo(0, 0);
    });
  });
  $("comboBtn").addEventListener("click", function () { $("comboPanel").hidden ? openCombo() : closeCombo(); });
  $("comboQ").addEventListener("input", renderCombo);
  $("comboList").addEventListener("click", function (e) {
    var li = e.target.closest("li[data-v]"); if (li) setFactory(li.getAttribute("data-v"));
  });
  document.addEventListener("click", function (e) { if (!$("combo").contains(e.target)) closeCombo(); });
  $("q").addEventListener("input", renderSearch);

  /* ---------- الفواتير ---------- */
  var INV = "majd-invoice-v1";
  var cart = [];            // [{f,m,p,q}]
  var pend = {};            // الكمية المختارة قبل الإضافة لكل موديل
  var invFac = "", invLimit = 40;
  function ikey(r) { return norm(r.f) + "|" + norm(r.m); }
  function loadInv() { try { cart = JSON.parse(localStorage.getItem(INV) || "[]"); } catch (e) { cart = []; } if (!Array.isArray(cart)) cart = []; }
  function saveInv() { try { localStorage.setItem(INV, JSON.stringify(cart)); } catch (e) {} }
  function num(p) { var n = parseFloat(String(p).replace(/,/g, "")); return isNaN(n) ? 0 : n; }
  function fmtN(n) { return n.toLocaleString("en-US", { maximumFractionDigits: 2 }); }
  function cartQty(k) { for (var i = 0; i < cart.length; i++) if (ikey(cart[i]) === k) return cart[i].q; return 0; }
  function clampQ(v) { v = parseInt(v, 10); return isNaN(v) || v < 1 ? 1 : Math.min(v, 99999); }

  function invPool() {
    var q = norm($("iq").value);
    var pool = invFac ? rows.filter(function (r) { return (r.f || "بدون مصنع") === invFac; }) : rows;
    if (q) pool = pool.filter(function (r) { return norm(r.m).indexOf(q) > -1 || norm(r.f).indexOf(q) > -1; });
    return { pool: pool, q: q };
  }
  function renderChips() {
    var h = '<button type="button" class="chip' + (invFac ? "" : " on") + '" data-f="">الكل</button>';
    h += facList.map(function (f) { return '<button type="button" class="chip' + (f === invFac ? " on" : "") + '" data-f="' + esc(f) + '">' + esc(f) + "</button>"; }).join("");
    $("chips").innerHTML = h;
    var on = $("chips").querySelector(".on"); if (on && on.scrollIntoView) on.scrollIntoView({ inline: "center", block: "nearest" });
  }
  function renderInvItems() {
    var box = $("ilist"), info = $("iinfo");
    if (!rows.length) { info.textContent = ""; box.innerHTML = '<div class="empty">لا توجد بيانات. افتح صفحة الإعدادات وحدّث من ملف إكسيل.</div>'; return; }
    var r = invPool(), pool = r.pool, q = r.q;
    if (!pool.length) { info.textContent = ""; box.innerHTML = '<div class="empty bad">غير موجود</div>'; return; }
    info.textContent = "عدد الموديلات: " + pool.length;
    var show = pool.slice(0, invLimit);
    box.innerHTML = show.map(function (x) {
      var k = ikey(x), n = pend[k] || 1, inC = cartQty(k);
      return '<div class="im" data-k="' + esc(k) + '">' +
        '<div class="im-top"><div class="im-info"><div class="im-f">' + hlText(x.f || "-", q) + '</div><div class="im-m">' + hlText(x.m, q) + '</div></div>' +
        '<div class="im-p">' + fmtPrice(x.p) + '</div></div>' +
        '<div class="im-act"><div class="qty"><button type="button" data-a="plus" aria-label="زيادة">+</button>' +
        '<input type="number" inputmode="numeric" min="1" value="' + n + '" aria-label="الكمية">' +
        '<button type="button" data-a="minus" aria-label="نقص">−</button></div>' +
        '<button type="button" class="btn" data-a="add">أضف للفاتورة</button></div>' +
        '<div class="in-cart"' + (inC ? "" : " hidden") + '>في الفاتورة: ' + inC + '</div></div>';
    }).join("") + (pool.length > invLimit ? '<button type="button" class="btn ghost more" data-a="more">عرض المزيد (' + (pool.length - invLimit) + ')</button>' : "");
  }
  function findRow(k) { for (var i = 0; i < rows.length; i++) if (ikey(rows[i]) === k) return rows[i]; return null; }

  function renderCart() {
    var box = $("cartList"), tot = 0, cnt = 0;
    cart.forEach(function (c) { tot += num(c.p) * c.q; cnt += c.q; });
    $("cartCount").textContent = cart.length;
    var has = cart.length > 0;
    $("cartTotal").hidden = !has; $("cartBtns").hidden = !has;
    if (!has) { box.innerHTML = '<div class="empty">الفاتورة فاضية. أضف موديلات من تبويب "الموديلات".</div>'; return; }
    box.innerHTML = cart.map(function (c, i) {
      return '<div class="im" data-i="' + i + '">' +
        '<div class="im-top"><div class="im-info"><div class="im-f">' + esc(c.f || "-") + '</div><div class="im-m">' + esc(c.m) + '</div></div>' +
        '<div class="im-p">' + fmtPrice(c.p) + '</div></div>' +
        '<div class="im-act"><div class="qty"><button type="button" data-a="cplus" aria-label="زيادة">+</button>' +
        '<input type="number" inputmode="numeric" min="1" value="' + c.q + '" aria-label="الكمية">' +
        '<button type="button" data-a="cminus" aria-label="نقص">−</button></div>' +
        '<div class="line-total">' + fmtN(num(c.p) * c.q) + '</div>' +
        '<button type="button" class="rm" data-a="rm">حذف</button></div></div>';
    }).join("");
    $("cartTotal").innerHTML = "<span>الإجمالي (" + cnt + " قطعة)</span><span>" + fmtN(tot) + "</span>";
  }
  function invRefresh() { renderChips(); renderInvItems(); renderCart(); }
  function invText() {
    var t = "فاتورة - مكتب المجد\n", tot = 0;
    cart.forEach(function (c, i) {
      var l = num(c.p) * c.q; tot += l;
      t += (i + 1) + ". " + (c.f ? c.f + " - " : "") + c.m + " | " + c.q + " × " + fmtN(num(c.p)) + " = " + fmtN(l) + "\n";
    });
    return t + "الإجمالي: " + fmtN(tot);
  }
  function setSeg(s) {
    $("segItems").classList.toggle("on", s === "items"); $("segCart").classList.toggle("on", s === "cart");
    $("segHist").classList.toggle("on", s === "hist");
    $("invItems").hidden = s !== "items"; $("invCart").hidden = s !== "cart"; $("invHist").hidden = s !== "hist";
    if (s !== "cart") $("closeBox").hidden = true;
    window.scrollTo(0, 0);
  }

  $("segItems").addEventListener("click", function () { setSeg("items"); });
  $("segCart").addEventListener("click", function () { setSeg("cart"); renderCart(); });
  $("segHist").addEventListener("click", function () { setSeg("hist"); renderHist(); });
  $("iq").addEventListener("input", function () { invLimit = 40; renderInvItems(); });
  $("chips").addEventListener("click", function (e) {
    var b = e.target.closest(".chip"); if (!b) return;
    invFac = b.getAttribute("data-f"); invLimit = 40; renderChips(); renderInvItems();
  });
  $("ilist").addEventListener("click", function (e) {
    var b = e.target.closest("[data-a]"); if (!b) return;
    var a = b.getAttribute("data-a");
    if (a === "more") { invLimit += 40; renderInvItems(); return; }
    var card = b.closest(".im"), k = card.getAttribute("data-k"), inp = card.querySelector("input");
    var n = clampQ(inp.value);
    if (a === "plus") n++; else if (a === "minus") n = Math.max(1, n - 1);
    if (a === "plus" || a === "minus") { inp.value = n; pend[k] = n; return; }
    if (a === "add") {
      var r = findRow(k); if (!r) return;
      var hit = null; cart.forEach(function (c) { if (ikey(c) === k) hit = c; });
      if (hit) { hit.q += n; hit.p = r.p; } else cart.push({ f: r.f, m: r.m, p: r.p, q: n });
      saveInv(); delete pend[k]; inp.value = 1;
      var ic = card.querySelector(".in-cart"); ic.textContent = "في الفاتورة: " + cartQty(k); ic.hidden = false;
      $("cartCount").textContent = cart.length;
      toast("اتضاف: " + r.m + " × " + n);
    }
  });
  $("ilist").addEventListener("input", function (e) {
    if (e.target.tagName !== "INPUT") return;
    var k = e.target.closest(".im").getAttribute("data-k"), v = parseInt(e.target.value, 10);
    if (!isNaN(v) && v > 0) pend[k] = v;
  });
  $("cartList").addEventListener("click", function (e) {
    var b = e.target.closest("[data-a]"); if (!b) return;
    var card = b.closest(".im"), i = +card.getAttribute("data-i"), c = cart[i]; if (!c) return;
    var a = b.getAttribute("data-a");
    if (a === "cplus") c.q = Math.min(99999, c.q + 1);
    else if (a === "cminus") c.q = Math.max(1, c.q - 1);
    else if (a === "rm") cart.splice(i, 1);
    saveInv(); renderCart(); renderInvItems();
  });
  $("cartList").addEventListener("change", function (e) {
    if (e.target.tagName !== "INPUT") return;
    var c = cart[+e.target.closest(".im").getAttribute("data-i")]; if (!c) return;
    c.q = clampQ(e.target.value); saveInv(); renderCart(); renderInvItems();
  });
  $("copyInv").addEventListener("click", function () {
    var t = invText();
    function fb() {
      var ta = document.createElement("textarea"); ta.value = t; document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); toast("تم نسخ الفاتورة."); } catch (e) { toast("تعذر النسخ."); }
      document.body.removeChild(ta);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(function () { toast("تم نسخ الفاتورة."); }, fb); else fb();
  });
  var invClearArmed = false;
  $("clearInv").addEventListener("click", function () {
    if (!invClearArmed) { invClearArmed = true; $("clearInvAsk").hidden = false; setTimeout(function () { invClearArmed = false; $("clearInvAsk").hidden = true; }, 5000); return; }
    cart = []; saveInv(); invClearArmed = false; $("clearInvAsk").hidden = true; renderCart(); renderInvItems(); toast("تم مسح الفاتورة.");
  });

  /* ---------- قفل الفاتورة + PDF + السجل ---------- */
  var HKEY = "majd-invoices-v1";
  var hist = [], pdfCache = {};
  function loadHist() { try { hist = JSON.parse(localStorage.getItem(HKEY) || "[]"); } catch (e) { hist = []; } if (!Array.isArray(hist)) hist = []; }
  function saveHist() { try { localStorage.setItem(HKEY, JSON.stringify(hist)); return true; } catch (e) { return false; } }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function dateStr(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function timeStr(d) { return pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  // رقم الفاتورة = اسم العميل + التاريخ (ولو نفس العميل في نفس اليوم يتضاف رقم تسلسلي)
  function makeNo(name, ds) {
    var n = 0; hist.forEach(function (h) { if (h.name === name && h.date === ds) n++; });
    return name + " - " + ds + (n ? " (" + (n + 1) + ")" : "");
  }
  function fileName(no) { return no.replace(/[\\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim() + ".pdf"; }
  function totals(items) { var t = 0, c = 0; items.forEach(function (x) { t += num(x.p) * x.q; c += x.q; }); return { total: t, count: c }; }

  /* رسم الفاتورة على canvas (المتصفح بيظبط الحروف العربية) ثم تحويلها لصفحات PDF */
  function loadImg(src) {
    return new Promise(function (res) { var i = new Image(); i.onload = function () { res(i); }; i.onerror = function () { res(null); }; i.src = src; });
  }
  function fontsReady() {
    if (!document.fonts || !document.fonts.load) return Promise.resolve();
    return Promise.race([
      Promise.all([document.fonts.load("800 30px Cairo"), document.fonts.load("600 30px Cairo")]),
      new Promise(function (r) { setTimeout(r, 2500); })
    ]).catch(function () {});
  }
  var PW = 1240, PH = 1754, MG = 70, FONT = 'Cairo,"Segoe UI",Tahoma,Arial,sans-serif';
  var COLS = [["#", 60], ["المصنع", 270], ["رقم الموديل", 250], ["الكمية", 110], ["السعر", 180], ["الإجمالي", 230]];
  function fit(ctx, text, maxW, size, weight) {
    var s = size; ctx.font = weight + " " + s + "px " + FONT;
    while (ctx.measureText(text).width > maxW && s > 14) { s -= 2; ctx.font = weight + " " + s + "px " + FONT; }
  }
  function cell(ctx, text, col, y, h, weight, color, ltr) {
    var right = PW - MG, i; for (i = 0; i < col; i++) right -= COLS[i][1];
    var w = COLS[col][1]; ctx.fillStyle = color; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    ctx.direction = ltr ? "ltr" : "rtl"; fit(ctx, String(text), w - 16, 28, weight);
    ctx.fillText(String(text), right - w / 2, y + h / 2 + 2);
  }
  function drawPages(inv, logo) {
    var RH = 62, TH = 60, items = inv.items, tt = totals(items);
    // توزيع الصفوف على الصفحات
    var pages = [], first = true, i = 0;
    while (i < items.length || first) {
      var top = first ? 410 : MG, cap = Math.floor((PH - 130 - top - TH) / RH), chunk = items.slice(i, i + cap);
      i += chunk.length; pages.push({ rows: chunk, top: top, start: i - chunk.length }); first = false;
      if (i >= items.length) { // مكان الإجمالي في آخر صفحة
        var used = top + TH + chunk.length * RH;
        if (used + 130 > PH - 100) pages.push({ rows: [], top: MG, start: i, onlyTotal: true });
        break;
      }
    }
    var out = [];
    pages.forEach(function (pg, pi) {
      var cv = document.createElement("canvas"); cv.width = PW; cv.height = PH;
      var ctx = cv.getContext("2d");
      ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, PW, PH);
      if (pi === 0) {
        var g = ctx.createLinearGradient(0, 0, PW, 0); g.addColorStop(0, "#2f86a8"); g.addColorStop(1, "#083a52");
        ctx.fillStyle = g; ctx.fillRect(0, 0, PW, 210);
        ctx.fillStyle = "#f2b705"; ctx.fillRect(0, 210, PW, 10);
        if (logo) {
          ctx.save(); ctx.beginPath(); ctx.arc(PW - MG - 65, 105, 65, 0, Math.PI * 2); ctx.clip();
          ctx.drawImage(logo, PW - MG - 130, 40, 130, 130); ctx.restore();
        }
        ctx.textBaseline = "alphabetic"; ctx.direction = "rtl"; ctx.textAlign = "right";
        ctx.fillStyle = "#fff"; ctx.font = "900 56px " + FONT; ctx.fillText("مكتب المجد", PW - MG - 160, 100);
        ctx.fillStyle = "#f2b705"; ctx.font = "700 32px " + FONT; ctx.fillText("لملابس الأطفال", PW - MG - 160, 150);
        ctx.textAlign = "left"; ctx.fillStyle = "#fff"; ctx.font = "900 60px " + FONT; ctx.fillText("فاتورة", MG, 125);
        // بيانات الفاتورة
        var y = 290; ctx.textAlign = "right";
        [["رقم الفاتورة", inv.no], ["العميل", inv.name], ["التاريخ", inv.date + " - " + inv.time]].forEach(function (r) {
          ctx.direction = "rtl"; ctx.fillStyle = "#4f6b79"; ctx.font = "700 30px " + FONT; ctx.fillText(r[0] + ":", PW - MG, y);
          var lw = ctx.measureText(r[0] + ":").width + 18;
          var vt = r[1].replace(/[0-9][0-9:\-]*[0-9]( \(\d+\))?/g, function (m) { return "\u2066" + m + "\u2069"; });
          ctx.fillStyle = "#12303f"; fit(ctx, vt, PW - 2 * MG - lw, 36, 800); ctx.fillText(vt, PW - MG - lw, y);
          y += 52;
        });
      }
      // رأس الجدول
      var ty = pg.top;
      if (!pg.onlyTotal) {
        ctx.fillStyle = "#0f5a78"; ctx.fillRect(MG, ty, PW - 2 * MG, TH);
        COLS.forEach(function (c, ci) { cell(ctx, c[0], ci, ty, TH, 800, "#fff", false); });
        pg.rows.forEach(function (r, ri) {
          var ry = ty + TH + ri * RH;
          ctx.fillStyle = ri % 2 ? "#eef6f9" : "#fff"; ctx.fillRect(MG, ry, PW - 2 * MG, RH);
          ctx.fillStyle = "#cfe3ea"; ctx.fillRect(MG, ry + RH - 1, PW - 2 * MG, 1);
          cell(ctx, pg.start + ri + 1, 0, ry, RH, 600, "#4f6b79", true);
          cell(ctx, r.f || "-", 1, ry, RH, 700, "#12303f", false);
          cell(ctx, r.m, 2, ry, RH, 800, "#0f5a78", true);
          cell(ctx, r.q, 3, ry, RH, 800, "#12303f", true);
          cell(ctx, fmtN(num(r.p)), 4, ry, RH, 700, "#12303f", true);
          cell(ctx, fmtN(num(r.p) * r.q), 5, ry, RH, 800, "#13804a", true);
        });
        ctx.strokeStyle = "#0f5a78"; ctx.lineWidth = 2; ctx.strokeRect(MG, ty, PW - 2 * MG, TH + pg.rows.length * RH);
        ty += TH + pg.rows.length * RH;
      }
      if (pi === pages.length - 1) {
        var by = ty + 30;
        ctx.fillStyle = "#0f5a78"; ctx.fillRect(MG, by, PW - 2 * MG, 90);
        ctx.textBaseline = "middle"; ctx.fillStyle = "#fff"; ctx.font = "800 34px " + FONT;
        ctx.direction = "rtl"; ctx.textAlign = "right"; ctx.fillText("الإجمالي (" + tt.count + " قطعة)", PW - MG - 24, by + 47);
        ctx.direction = "ltr"; ctx.textAlign = "left"; ctx.fillStyle = "#f2b705"; ctx.font = "900 44px " + FONT; ctx.fillText(fmtN(tt.total), MG + 24, by + 47);
      }
      ctx.textBaseline = "alphabetic"; ctx.textAlign = "center"; ctx.direction = "rtl"; ctx.fillStyle = "#4f6b79"; ctx.font = "600 24px " + FONT;
      ctx.fillText("مكتب المجد لملابس الأطفال  •  صفحة " + (pi + 1) + " من " + pages.length, PW / 2, PH - 50);
      out.push(cv);
    });
    return out;
  }
  function u8(str) { var b = new Uint8Array(str.length); for (var i = 0; i < str.length; i++) b[i] = str.charCodeAt(i) & 255; return b; }
  function dataToBytes(url) { var bin = atob(url.split(",")[1]), b = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); return b; }
  function buildPdf(canvases) {
    var parts = [], offs = [], len = 0, n = canvases.length;
    function push(x) { var b = typeof x === "string" ? u8(x) : x; parts.push(b); len += b.length; }
    function obj(id, body, stream) {
      offs[id] = len; push(id + " 0 obj\n" + body);
      if (stream) { push("\nstream\n"); push(stream); push("\nendstream"); }
      push("\nendobj\n");
    }
    push("%PDF-1.4\n");
    var kids = []; for (var i = 0; i < n; i++) kids.push((3 + 3 * i) + " 0 R");
    obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
    obj(2, "<< /Type /Pages /Kids [" + kids.join(" ") + "] /Count " + n + " >>");
    canvases.forEach(function (cv, i) {
      var p = 3 + 3 * i, jpg = dataToBytes(cv.toDataURL("image/jpeg", 0.9)), content = "q 595.28 0 0 841.89 0 0 cm /Im0 Do Q";
      obj(p, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Im0 " + (p + 2) + " 0 R >> >> /Contents " + (p + 1) + " 0 R >>");
      obj(p + 1, "<< /Length " + content.length + " >>", u8(content));
      obj(p + 2, "<< /Type /XObject /Subtype /Image /Width " + cv.width + " /Height " + cv.height + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpg.length + " >>", jpg);
    });
    var xr = len, total = 3 + 3 * n, s = "xref\n0 " + total + "\n0000000000 65535 f \n";
    for (var k = 1; k < total; k++) s += ("0000000000" + offs[k]).slice(-10) + " 00000 n \n";
    push(s + "trailer\n<< /Size " + total + " /Root 1 0 R >>\nstartxref\n" + xr + "\n%%EOF");
    return new Blob(parts, { type: "application/pdf" });
  }
  function makePdf(inv) {
    return Promise.all([fontsReady(), loadImg("icons/logo-header.jpg")]).then(function (r) {
      var blob = buildPdf(drawPages(inv, r[1]));
      return new File([blob], fileName(inv.no), { type: "application/pdf" });
    });
  }
  function getPdf(inv) { return pdfCache[inv.id] ? Promise.resolve(pdfCache[inv.id]) : makePdf(inv).then(function (f) { pdfCache[inv.id] = f; return f; }); }
  function download(file) {
    var u = URL.createObjectURL(file), a = document.createElement("a");
    a.href = u; a.download = file.name; document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(u); }, 1500);
  }
  function shareInv(inv) {
    toast("جارٍ تجهيز الفاتورة…");
    getPdf(inv).then(function (file) {
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        return navigator.share({ files: [file], title: inv.no }).then(function () { $("msg").hidden = true; });
      }
      download(file); toast("تم تحميل الـ PDF. ابعته من الواتس من التحميلات.");
    }).catch(function (e) {
      if (e && e.name === "AbortError") { $("msg").hidden = true; return; }
      getPdf(inv).then(function (f) { download(f); toast("تعذرت المشاركة المباشرة، اتحمّل الملف. ابعته من الواتس."); }).catch(function () { toast("تعذر إنشاء الـ PDF."); });
    });
  }
  function downloadInv(inv) {
    getPdf(inv).then(function (f) { download(f); toast("تم تحميل الملف."); }).catch(function () { toast("تعذر إنشاء الـ PDF."); });
  }

  function renderHist() {
    var box = $("histList");
    if (!hist.length) { box.innerHTML = '<div class="empty">لا توجد فواتير مقفولة. اقفل فاتورة وهتظهر هنا.</div>'; return; }
    box.innerHTML = hist.map(function (h) {
      return '<div class="im" data-id="' + esc(h.id) + '"><div class="im-top"><div class="im-info"><div class="im-f">' + esc(h.date + "  " + h.time) + ' • ' + h.items.length + ' صنف</div>' +
        '<div class="im-no">' + esc(h.no) + '</div></div><div class="im-p">' + fmtN(h.total) + '</div></div>' +
        '<div class="im-act"><button type="button" class="btn" data-a="share">مشاركة PDF</button>' +
        '<button type="button" class="btn ghost" data-a="dl">تحميل</button>' +
        '<button type="button" class="rm" data-a="del">حذف</button></div></div>';
    }).join("");
  }
  var delArmed = "";
  $("histList").addEventListener("click", function (e) {
    var b = e.target.closest("[data-a]"); if (!b) return;
    var id = b.closest(".im").getAttribute("data-id"), inv = null, idx = -1;
    hist.forEach(function (h, i) { if (h.id === id) { inv = h; idx = i; } });
    if (!inv) return;
    var a = b.getAttribute("data-a");
    if (a === "share") shareInv(inv);
    else if (a === "dl") downloadInv(inv);
    else if (a === "del") {
      if (delArmed !== id) { delArmed = id; b.textContent = "تأكيد الحذف"; setTimeout(function () { delArmed = ""; if (b.isConnected) b.textContent = "حذف"; }, 4000); return; }
      hist.splice(idx, 1); delete pdfCache[id]; saveHist(); delArmed = ""; renderHist(); toast("تم حذف الفاتورة.");
    }
  });

  $("closeInv").addEventListener("click", function () {
    if (!cart.length) return;
    $("closeBox").hidden = false; $("closeInv").parentNode.hidden = true;
    $("custName").value = ""; updNo(); $("custName").focus();
    $("closeBox").scrollIntoView({ block: "center" });
  });
  function updNo() {
    var nm = $("custName").value.trim();
    $("closeNo").textContent = "رقم الفاتورة: " + (nm ? makeNo(nm, dateStr(new Date())) : "اسم العميل - التاريخ");
  }
  $("custName").addEventListener("input", updNo);
  $("cancelClose").addEventListener("click", function () { $("closeBox").hidden = true; $("cartBtns").hidden = !cart.length; });
  var closing = false;
  function confirmClose() {
    var nm = $("custName").value.trim().replace(/\s+/g, " ");
    if (!nm) { toast("اكتب اسم العميل الأول."); $("custName").focus(); return; }
    if (!cart.length || closing) return;
    closing = true;
    var d = new Date(), ds = dateStr(d), items = cart.map(function (c) { return { f: c.f, m: c.m, p: c.p, q: c.q }; });
    var inv = { id: String(d.getTime()), name: nm, date: ds, time: timeStr(d), no: makeNo(nm, ds), items: items, total: totals(items).total };
    toast("جارٍ إنشاء الـ PDF…");
    makePdf(inv).then(function (f) {
      pdfCache[inv.id] = f; hist.unshift(inv); saveHist();
      cart = []; saveInv(); $("closeBox").hidden = true;
      renderCart(); renderInvItems(); setSeg("hist"); renderHist();
      toast("تم قفل الفاتورة. اضغط \"مشاركة PDF\" لإرسالها.");
    }).catch(function () { toast("تعذر إنشاء الـ PDF. الفاتورة لسه مفتوحة."); })
      .then(function () { closing = false; });
  }
  $("confirmClose").addEventListener("click", confirmClose);
  $("custName").addEventListener("keydown", function (e) { if (e.key === "Enter") confirmClose(); });

  /* ---------- تشغيل ---------- */
  load(); loadInv(); loadHist(); renderHist(); fillFactories(); renderSearch(); renderStat(); invRefresh();
  // أول مرة بدون بيانات: حاول تحمّل data.xlsx لو موجود مع التطبيق
  if (!rows.length) {
    fetch("data.xlsx", { cache: "no-store" }).then(function (r) { if (!r.ok) throw 0; return r.arrayBuffer(); })
      .then(function (b) { applyRows(parseWorkbookBuffer(b), "data.xlsx"); }).catch(function () {});
  }
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(function () {});
  window.__majd = { parseMatrix: parseMatrix };
})();
