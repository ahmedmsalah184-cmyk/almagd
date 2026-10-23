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
    $("invItems").hidden = s !== "items"; $("invCart").hidden = s !== "cart";
  }

  $("segItems").addEventListener("click", function () { setSeg("items"); });
  $("segCart").addEventListener("click", function () { setSeg("cart"); renderCart(); });
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
  $("shareWa").addEventListener("click", function () { window.open("https://wa.me/?text=" + encodeURIComponent(invText()), "_blank"); });
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

  /* ---------- تشغيل ---------- */
  load(); loadInv(); fillFactories(); renderSearch(); renderStat(); invRefresh();
  // أول مرة بدون بيانات: حاول تحمّل data.xlsx لو موجود مع التطبيق
  if (!rows.length) {
    fetch("data.xlsx", { cache: "no-store" }).then(function (r) { if (!r.ok) throw 0; return r.arrayBuffer(); })
      .then(function (b) { applyRows(parseWorkbookBuffer(b), "data.xlsx"); }).catch(function () {});
  }
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(function () {});
  window.__majd = { parseMatrix: parseMatrix };
})();
