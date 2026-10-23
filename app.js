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
    fillFactories(); renderSearch(); renderStat();
    toast(ok ? "تم تحديث البيانات: " + n.length + " موديل." : "اتحدثت البيانات لكن مقدرتش أحفظها على الموبايل.");
  }

  /* ---------- صفحة البحث ---------- */
  function factories() {
    var s = {}, out = [];
    rows.forEach(function (r) { var k = r.f || "بدون مصنع"; if (!s[k]) { s[k] = 1; out.push(k); } });
    return out.sort(function (a, b) { return a.localeCompare(b, "ar"); });
  }
  function fillFactories() {
    var sel = $("factory"), cur = sel.value;
    var list = factories();
    sel.innerHTML = '<option value="">كل المصانع (' + list.length + ")</option>" +
      list.map(function (f) { return '<option value="' + esc(f) + '">' + esc(f) + "</option>"; }).join("");
    if (list.indexOf(cur) > -1) sel.value = cur;
  }
  function hl(model, q) {
    var raw = String(model), j = raw.toUpperCase().indexOf(q);
    if (!q || j < 0) return esc(raw);
    return esc(raw.slice(0, j)) + "<mark>" + esc(raw.slice(j, j + q.length)) + "</mark>" + esc(raw.slice(j + q.length));
  }
  function card(r, q) {
    return '<div class="hit"><div class="m">' + hl(r.m, q) + '</div>' +
      '<div class="row"><span>اسم المصنع</span><strong>' + esc(r.f || "-") + '</strong></div>' +
      '<div class="row"><span>السعر</span><span class="price">' + fmtPrice(r.p) + "</span></div></div>";
  }
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
      info.textContent = fac ? "عدد موديلات المصنع: " + pool.length : "اختار مصنع أو اكتب رقم الموديل";
      if (!fac) { out.innerHTML = ""; return; }
      out.innerHTML = pool.slice(0, 60).map(function (r) { return card(r, ""); }).join("") +
        (pool.length > 60 ? '<div class="info">أول 60 موديل. اكتب رقم الموديل لتضييق البحث.</div>' : "");
      return;
    }
    var exact = pool.filter(function (r) { return norm(r.m) === q; });
    var list = exact.length ? exact : pool.filter(function (r) { return norm(r.m).indexOf(q) > -1; });
    if (!list.length) { info.textContent = ""; out.innerHTML = '<div class="empty bad">غير موجود</div>'; return; }
    info.textContent = exact.length ? "" : "نتائج مشابهة (" + list.length + ")";
    out.innerHTML = list.slice(0, 50).map(function (r) { return card(r, q); }).join("");
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
    rows = []; meta = {}; save(); fillFactories(); renderSearch(); renderStat();
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
      $("view-settings").hidden = b.dataset.view !== "settings";
      window.scrollTo(0, 0);
    });
  });
  $("factory").addEventListener("change", renderSearch);
  $("q").addEventListener("input", renderSearch);

  /* ---------- تشغيل ---------- */
  load(); fillFactories(); renderSearch(); renderStat();
  // أول مرة بدون بيانات: حاول تحمّل data.xlsx لو موجود مع التطبيق
  if (!rows.length) {
    fetch("data.xlsx", { cache: "no-store" }).then(function (r) { if (!r.ok) throw 0; return r.arrayBuffer(); })
      .then(function (b) { applyRows(parseWorkbookBuffer(b), "data.xlsx"); }).catch(function () {});
  }
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(function () {});
  window.__majd = { parseMatrix: parseMatrix };
})();
