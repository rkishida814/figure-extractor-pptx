"use strict";

const $ = (id) => document.getElementById(id);
const EXTS = ["png", "tif", "bmp", "jpg"];

let data = null;          // /api/upload の結果
let cur = 0;              // 表示中のスライド
let selected = [];        // 図の一覧（選んだ順）[{key, slide, ids:[並び順...]}]。ids が複数なら1枚にまとめる
let custom = {};          // 図の key -> {name?, ext?}（選び直しても覚えておく）
let gseq = 0;             // 図の key の通し番号
let history = [];         // 選択の取り消し用
let defaultExt = "png";
let previews = {};        // 図の中身（sig）-> {url} / {pending} / {error}
let hoverKey = null;

const keyOf = (slide, id) => `${slide}:${id}`;
const shapeOf = (slide, id) => data.slides[slide].shapes.find((s) => s.id === id);
const entryIndexOf = (slide, id) => selected.findIndex((e) => e.slide === slide && e.ids.includes(id));

// ------------------------------------------------------------------ 読み込み

async function init() {
  try {
    const st = await (await fetch("/api/status")).json();
    // 前回の設定（アプリ側のファイルに保存している）
    const set = st.settings || {};
    $("folder").value = set.folder || "";
    $("dpi").value = String(set.dpi || 300);
    $("transparent").checked = set.transparent === undefined ? true : !!set.transparent;
    if (set.ext && EXTS.includes(set.ext)) defaultExt = set.ext;
    $("rendererInfo").textContent = st.renderer
      ? `描画に使うアプリ: ${st.renderer === "powerpoint" ? "Microsoft PowerPoint" : "LibreOffice"}`
      : "PowerPoint も LibreOffice も見つかりません。どちらかを入れてください。";
    if (!$("folder").value) $("folder").placeholder = st.home + "/Desktop/figures";
  } catch { /* なし */ }
}

async function loadFile(file) {
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  busy(`「${file.name}」を読み込んでいます…（PowerPoint で PDF にしています）`);
  try {
    const r = await fetch("/api/upload", { method: "POST", body: fd });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error || r.statusText);
    data = j;
    cur = 0; selected = []; custom = {}; history = []; previews = {}; currentKey = null; cycleKey = null; lastPick = null;
    $("fileName").textContent = j.filename;
    $("status").textContent = `${j.slides.length} 枚`;
    $("drop").hidden = true;
    $("stage").hidden = false;
    $("slideNav").hidden = false;
    $("result").textContent = "";
    buildThumbs();
    showSlide(0);
    renderList();
  } catch (e) {
    alert("読み込みに失敗しました\n" + e.message);
  } finally {
    busy(null);
  }
}

function busy(msg) {
  let el = document.querySelector(".busy");
  if (!msg) { if (el) el.remove(); return; }
  if (!el) { el = document.createElement("div"); el.className = "busy"; el.appendChild(document.createElement("div")); document.body.appendChild(el); }
  el.firstChild.textContent = msg;
}

// ------------------------------------------------------------------ スライド表示

function buildThumbs() {
  const nav = $("thumbs");
  nav.innerHTML = "";
  data.slides.forEach((s, i) => {
    const d = document.createElement("div");
    d.className = "thumb" + (s.hidden ? " hiddenSlide" : "");
    d.innerHTML = `<img src="${s.image}" alt=""><span class="no">${i + 1}${s.hidden ? "（非表示）" : ""}</span><span class="badge" hidden></span>`;
    d.onclick = () => showSlide(i);
    nav.appendChild(d);
  });
}

function showSlide(i) {
  if (!data) return;
  cur = Math.max(0, Math.min(data.slides.length - 1, i));
  const img = $("slideImg");
  img.src = data.slides[cur].image;
  $("slideLabel").textContent = `スライド ${cur + 1} / ${data.slides.length}`;
  [...$("thumbs").children].forEach((t, k) => t.classList.toggle("active", k === cur));
  $("thumbs").children[cur]?.scrollIntoView({ block: "nearest" });
  fitStage();
  drawOverlay();
}

function fitStage() {
  if (!data) return;
  const wrap = $("stageWrap");
  const maxW = wrap.clientWidth - 32;
  const maxH = wrap.clientHeight - 80;
  const ratio = data.slideW / data.slideH;
  const w = Math.max(200, Math.min(maxW, maxH * ratio));
  $("slideImg").style.width = w + "px";
  drawOverlay();
}

function scale() { return $("slideImg").clientWidth / data.slideW; }

function drawOverlay() {
  if (!data) return;
  const ov = $("overlay");
  ov.innerHTML = "";
  const s = scale();
  const order = new Map();
  selected.forEach((e, k) => { if (e.slide === cur) e.ids.forEach((id) => order.set(id, k + 1)); });
  // まとめた図は、全体を囲む破線の枠も描く
  selected.forEach((e, k) => {
    if (e.slide !== cur || e.ids.length < 2) return;
    const b = entryBox(e);
    const g = document.createElement("div");
    g.className = "box grp" + (e.key === currentKey ? " current" : "");
    Object.assign(g.style, { left: b.x * s - 4 + "px", top: b.y * s - 4 + "px", width: b.w * s + 8 + "px", height: b.h * s + 8 + "px" });
    const t = document.createElement("span"); t.className = "num"; t.textContent = `${k + 1}（${e.ids.length}個）`;
    g.appendChild(t);
    ov.appendChild(g);
  });
  for (const sh of data.slides[cur].shapes) {
    const key = keyOf(cur, sh.id);
    const n = order.get(sh.id);
    const k = n ? n - 1 : -1;
    if (!n && key !== hoverKey) continue;
    const b = document.createElement("div");
    b.className = "box " + (n ? "sel" : "hover") + (n && selected[k].key === currentKey ? " current" : "");
    Object.assign(b.style, {
      left: sh.x * s + "px", top: sh.y * s + "px",
      width: Math.max(2, sh.w * s) + "px", height: Math.max(2, sh.h * s) + "px",
      transform: sh.rot ? `rotate(${sh.rot}deg)` : "",
    });
    if (n && selected[k].ids.length === 1) { const t = document.createElement("span"); t.className = "num"; t.textContent = n; b.appendChild(t); }
    ov.appendChild(b);
  }
  // サムネイルの選択数
  const counts = {};
  selected.forEach((it) => { counts[it.slide] = (counts[it.slide] || 0) + 1; });
  [...$("thumbs").children].forEach((t, k) => {
    const badge = t.querySelector(".badge");
    badge.hidden = !counts[k];
    badge.textContent = counts[k] || "";
  });
}

// ------------------------------------------------------------------ 当たり判定

function toEmu(ev) {
  const r = $("slideImg").getBoundingClientRect();
  const s = scale();
  return { x: (ev.clientX - r.left) / s, y: (ev.clientY - r.top) / s };
}

// その点にあるオブジェクトを前面から順に返す
function hitAll(p) {
  const tol = 5 / scale();
  const shapes = data.slides[cur].shapes;
  const out = [];
  for (let k = shapes.length - 1; k >= 0; k--) {   // 後ろほど前面
    const sh = shapes[k];
    const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
    const a = (-(sh.rot || 0) * Math.PI) / 180;
    const dx = p.x - cx, dy = p.y - cy;
    const rx = dx * Math.cos(a) - dy * Math.sin(a);
    const ry = dx * Math.sin(a) + dy * Math.cos(a);
    if (Math.abs(rx) <= sh.w / 2 + tol && Math.abs(ry) <= sh.h / 2 + tol) out.push(sh);
  }
  return out;
}

function hit(p) { return hitAll(p)[0] || null; }

// 同じ場所をもう一度クリック（または ⌥+クリック）したら、1つ下のオブジェクトを返す
let lastPick = null;   // {slide, x, y, id}
function pickAt(p, cycle) {
  const hits = hitAll(p);
  if (!hits.length) { lastPick = null; return null; }
  let target = hits[0];
  const tol = 4 / scale();
  const same = lastPick && lastPick.slide === cur && Math.abs(lastPick.x - p.x) <= tol && Math.abs(lastPick.y - p.y) <= tol;
  if (cycle || same) {
    const k = lastPick && lastPick.slide === cur ? hits.findIndex((h) => h.id === lastPick.id) : -1;
    if (k >= 0 && (same || cycle)) target = hits[(k + 1) % hits.length];
    else if (cycle && hits.length > 1 && !same) target = hits[1];
  }
  lastPick = { slide: cur, x: p.x, y: p.y, id: target.id };
  return target;
}

function aabb(sh) {
  const a = ((sh.rot || 0) * Math.PI) / 180;
  const c = Math.abs(Math.cos(a)), s = Math.abs(Math.sin(a));
  const w = sh.w * c + sh.h * s, h = sh.w * s + sh.h * c;
  const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

// ------------------------------------------------------------------ 選択
//
// 一覧の 1 行が 1 枚の図。
// - クリック：新しい図を始める（すでにどこかの図に入っているオブジェクトなら、その図を「いまの図」にする）
// - ⌘・Ctrl・Shift+クリック：いまの図に足す（入っていれば外す）→ 複数のオブジェクトが 1 枚になる
// - ドラッグ：囲んだオブジェクトで新しい図を 1 枚作る（⌘ 等を押しながらなら、いまの図に足す）
// - 同じ場所をもう一度クリック（⌥+クリック）：その図の中身を、重なった下のオブジェクトに替える

let currentKey = null;     // いまの図

function entryBox(e) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  e.ids.forEach((id) => {
    const b = aabb(shapeOf(e.slide, id));
    x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h);
  });
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

function pushHistory() {
  history.push({ list: selected.map((e) => ({ ...e, ids: [...e.ids] })), current: currentKey });
  if (history.length > 100) history.shift();
}

function commit(list, current) {
  pushHistory();
  selected = list.filter((e) => e.ids.length > 0);
  currentKey = selected.some((e) => e.key === current) ? current : null;
  drawOverlay();
  renderList();
}

function setCurrent(key) {
  currentKey = key;
  drawOverlay();
  renderList();
}

function newEntry(slide, ids) { return { key: `f${++gseq}`, slide, ids: [...ids].sort((a, b) => a - b) }; }

// ids をほかの図から抜いた一覧を返す
function without(list, slide, ids) {
  return list.map((e) => (e.slide === slide ? { ...e, ids: e.ids.filter((x) => !ids.includes(x)) } : e));
}

function startFigure(ids) {
  if (!ids.length) return;
  const e = newEntry(cur, ids);
  commit([...without(selected, cur, ids), e], e.key);
}

function addToCurrent(ids) {
  const curE = selected.find((e) => e.key === currentKey);
  if (!curE || curE.slide !== cur) { startFigure(ids); return; }   // 別のスライドの図には足せない
  const list = without(selected, cur, ids).map((e) =>
    e.key === curE.key ? { ...e, ids: [...new Set([...curE.ids, ...ids])].sort((a, b) => a - b) } : e);
  commit(list, curE.key);
}

function toggleInCurrent(id) {
  const curE = selected.find((e) => e.key === currentKey);
  if (curE && curE.slide === cur && curE.ids.includes(id)) {
    commit(selected.map((e) => (e.key === curE.key ? { ...e, ids: e.ids.filter((x) => x !== id) } : e)), curE.key);
  } else {
    addToCurrent([id]);
  }
}

function addMany(shapes) { addToCurrent(shapes.map((sh) => sh.id)); }

let drag = null;
let cycleKey = null;   // 同じ場所の再クリックで中身を替える図

// スライドの外（周りの灰色の部分）から始めたドラッグも範囲選択にする
$("stageWrap").addEventListener("mousedown", (ev) => {
  if (ev.button !== 0 || !data || $("stage").hidden) return;
  if (ev.target.closest("#slideNav, button, input, select, a")) return;
  ev.preventDefault();
  drag = { start: toEmu(ev), sx: ev.clientX, sy: ev.clientY, moved: false, add: ev.metaKey || ev.ctrlKey || ev.shiftKey, alt: ev.altKey };
});

window.addEventListener("mousemove", (ev) => {
  if (!data) return;
  if (drag) {
    if (!drag.moved && Math.hypot(ev.clientX - drag.sx, ev.clientY - drag.sy) > 4) drag.moved = true;
    if (drag.moved) {
      const p = toEmu(ev), s = scale();
      const band = $("band");
      band.hidden = false;
      Object.assign(band.style, {
        left: Math.min(p.x, drag.start.x) * s + "px", top: Math.min(p.y, drag.start.y) * s + "px",
        width: Math.abs(p.x - drag.start.x) * s + "px", height: Math.abs(p.y - drag.start.y) * s + "px",
      });
    }
    return;
  }
  if (ev.target.closest && ev.target.closest("#stage")) {
    const sh = hit(toEmu(ev));
    const k = sh ? keyOf(cur, sh.id) : null;
    $("stage").style.cursor = sh ? "pointer" : "default";
    if (k !== hoverKey) { hoverKey = k; drawOverlay(); highlightRow(k); }
  } else if (hoverKey) { hoverKey = null; drawOverlay(); highlightRow(null); }
});

window.addEventListener("mouseup", (ev) => {
  if (!drag) return;
  const d = drag;
  drag = null;
  $("band").hidden = true;
  if (d.moved) {
    cycleKey = null; lastPick = null;
    const p = toEmu(ev);
    const x0 = Math.min(p.x, d.start.x), x1 = Math.max(p.x, d.start.x);
    const y0 = Math.min(p.y, d.start.y), y1 = Math.max(p.y, d.start.y);
    const ids = data.slides[cur].shapes.filter((sh) => {
      const b = aabb(sh);
      return b.x >= x0 && b.y >= y0 && b.x + b.w <= x1 && b.y + b.h <= y1;
    }).map((sh) => sh.id);
    if (d.add) addToCurrent(ids); else startFigure(ids);
    return;
  }

  const prev = lastPick ? { ...lastPick } : null;
  const sh = pickAt(d.start, d.alt);
  if (!sh) { cycleKey = null; if (currentKey) setCurrent(null); return; }
  const cycled = prev && prev.slide === cur && prev.id !== sh.id && hitAll(d.start).some((h) => h.id === prev.id);

  if (d.add) {
    if (cycled && currentKey) {   // ⌘+再クリック：いま足したものを下のオブジェクトに替える
      const list = selected.map((e) => (e.key === currentKey ? { ...e, ids: e.ids.filter((x) => x !== prev.id) } : e));
      selected = list;
    }
    if (d.alt || cycled) addToCurrent([sh.id]); else toggleInCurrent(sh.id);
    return;
  }

  // 再クリック：直前のクリックで作った図の中身を、下のオブジェクトに替える
  if (cycled && cycleKey) {
    const e = selected.find((x) => x.key === cycleKey);
    if (e && e.ids.length === 1 && e.ids[0] === prev.id) {
      const list = without(selected, cur, [sh.id]).map((x) => (x.key === e.key ? { ...x, ids: [sh.id] } : x));
      commit(list, e.key);
      return;
    }
  }
  const k = entryIndexOf(cur, sh.id);
  if (k >= 0) { cycleKey = selected[k].ids.length === 1 ? selected[k].key : null; setCurrent(selected[k].key); return; }
  startFigure([sh.id]);
  cycleKey = currentKey;
});

// ------------------------------------------------------------------ 一覧

function nameOf(it, k) { return custom[it.key]?.name ?? `fig${k + 1}`; }
function extOf(it) { return custom[it.key]?.ext ?? defaultExt; }

function renderList() {
  const ol = $("list");
  ol.innerHTML = "";
  $("count").textContent = selected.length;
  $("emptyMsg").hidden = selected.length > 0;
  $("saveBtn").disabled = selected.length === 0;
  selected.forEach((it, k) => {
    const li = document.createElement("li");
    li.dataset.key = it.key;
    if (it.ids.length > 1) li.classList.add("merged");
    if (it.key === currentKey) li.classList.add("current");
    li.addEventListener("click", (e) => {
      if (e.target.closest("input, select, button")) return;
      if (it.slide !== cur) showSlide(it.slide);
      setCurrent(it.key);
    });

    const idx = document.createElement("span");
    idx.className = "idx"; idx.textContent = k + 1;

    // 実際に抽出した画像（背景透明）を表示する
    const crop = document.createElement("div");
    crop.className = "crop";
    const pv = previews[sigOf(it)];
    if (pv && pv.url) {
      const img = document.createElement("img");
      img.src = pv.url; img.alt = ""; img.draggable = false;
      crop.appendChild(img);
    } else {
      crop.classList.add("wait");
      crop.textContent = pv && pv.error ? "抽出できません" : "抽出中…";
      if (pv && pv.error) crop.title = pv.error;
    }

    const fields = document.createElement("div");
    fields.className = "fields";
    const row = document.createElement("div");
    row.className = "nameRow";
    const name = document.createElement("input");
    name.className = "name" + (custom[it.key]?.name !== undefined ? " custom" : "");
    name.value = nameOf(it, k);
    name.spellcheck = false;
    name.addEventListener("input", () => {
      const def = `fig${selected.findIndex((x) => x.key === it.key) + 1}`;
      custom[it.key] = custom[it.key] || {};
      if (name.value === def) delete custom[it.key].name; else custom[it.key].name = name.value;
      name.classList.toggle("custom", custom[it.key].name !== undefined);
    });
    name.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || (e.key === "ArrowDown")) { e.preventDefault(); ol.children[k + 1]?.querySelector("input.name")?.focus(); }
      if (e.key === "ArrowUp") { e.preventDefault(); ol.children[k - 1]?.querySelector("input.name")?.focus(); }
    });
    const ext = document.createElement("select");
    ext.className = "ext";
    EXTS.forEach((e) => { const o = document.createElement("option"); o.value = e; o.textContent = "." + e; ext.appendChild(o); });
    ext.value = extOf(it);
    ext.onchange = () => { custom[it.key] = custom[it.key] || {}; custom[it.key].ext = ext.value; };
    row.append(name, ext);
    const src = document.createElement("div");
    src.className = "src";
    if (it.ids.length === 1) {
      const sh = shapeOf(it.slide, it.ids[0]);
      src.textContent = `スライド${it.slide + 1}・${sh.kind}・${sh.name}`;
    } else {
      src.textContent = `スライド${it.slide + 1}・${it.ids.length}個のオブジェクト`;
      src.title = it.ids.map((id) => shapeOf(it.slide, id).name).join("\n");
    }
    fields.append(row, src);

    const acts = document.createElement("div");
    acts.className = "acts";
    const rm = document.createElement("button");
    rm.className = "rm"; rm.textContent = "×"; rm.title = "一覧から外す";
    rm.onclick = () => commit(selected.filter((x) => x.key !== it.key), currentKey);
    acts.appendChild(rm);

    li.append(idx, crop, fields, acts);
    li.addEventListener("mouseenter", () => { if (it.slide === cur) { hoverKey = keyOf(cur, it.ids[0]); drawOverlay(); } });
    li.addEventListener("mouseleave", () => { hoverKey = null; drawOverlay(); });
    ol.appendChild(li);
  });
  if (selected.some((e) => !previews[sigOf(e)])) requestPreviews();
}

function sigOf(e) { return `${e.slide}:${[...e.ids].sort((a, b) => a - b).join(",")}`; }

let previewTimer = null;
function requestPreviews() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(async () => {
    if (!data) return;
    const session = data.session;
    const need = selected.filter((e) => !previews[sigOf(e)]);
    if (!need.length) return;
    need.forEach((e) => { previews[sigOf(e)] = { pending: true }; });
    const keyToSig = Object.fromEntries(need.map((e) => [e.key, sigOf(e)]));
    try {
      const r = await fetch("/api/preview", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session, items: need.map((e) => ({ key: e.key, slide: e.slide, shapes: e.ids })) }),
      });
      const j = await r.json();
      if (!data || data.session !== session) return;
      if (!r.ok) throw new Error(j.error || r.statusText);
      for (const [key, v] of Object.entries(j.previews)) previews[keyToSig[key]] = v;
    } catch (err) {
      Object.values(keyToSig).forEach((sg) => { if (previews[sg]?.pending) previews[sg] = { error: err.message }; });
    }
    renderList();
  }, 250);
}

function highlightRow(shapeKey) {
  let ek = null;
  if (shapeKey) { const [sl, id] = shapeKey.split(":").map(Number); const k = entryIndexOf(sl, id); if (k >= 0) ek = selected[k].key; }
  [...$("list").children].forEach((li) => li.classList.toggle("hl", li.dataset.key === ek));
}

$("bulkExt").onchange = (e) => {
  const v = e.target.value;
  if (!v) return;
  defaultExt = v;
  saveSettings({ ext: v });
  selected.forEach((it) => { custom[it.key] = custom[it.key] || {}; custom[it.key].ext = v; });
  e.target.value = "";
  renderList();
};

$("clearBtn").onclick = () => commit([], null);
$("selAllBtn").onclick = () => addMany(data.slides[cur].shapes);
$("prevBtn").onclick = () => showSlide(cur - 1);
$("nextBtn").onclick = () => showSlide(cur + 1);

// ------------------------------------------------------------------ 保存

$("chooseBtn").onclick = async () => {
  const r = await fetch("/api/choose-folder", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ start: $("folder").value }),
  });
  const j = await r.json();
  if (j.folder) { $("folder").value = j.folder; saveSettings({ folder: j.folder }); }
};
function saveSettings(obj) {
  fetch("/api/settings", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(obj) }).catch(() => {});
}
$("folder").onchange = () => saveSettings({ folder: $("folder").value.trim() });
$("dpi").onchange = () => saveSettings({ dpi: Number($("dpi").value) });
$("transparent").onchange = () => saveSettings({ transparent: $("transparent").checked });

$("saveBtn").onclick = () => doExport(false);

async function doExport(overwrite) {
  if (!data || !selected.length) return;
  const folder = $("folder").value.trim() || $("folder").placeholder;
  $("folder").value = folder;
  const items = selected.map((it, k) => ({ slide: it.slide, shapes: it.ids, name: nameOf(it, k).trim(), ext: extOf(it) }));
  busy(`${items.length} 個の図を書き出しています…`);
  let r, j;
  try {
    r = await fetch("/api/export", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        session: data.session, folder, items, overwrite,
        dpi: Number($("dpi").value), transparent: $("transparent").checked,
      }),
    });
    j = await r.json();
  } catch (e) {
    busy(null); showResult(`<span class="ng">${esc(e.message)}</span>`); return;
  }
  busy(null);
  if (r.status === 409) {
    if (confirm(`次のファイルがすでにあります。上書きしますか？\n\n${j.conflict.join("\n")}`)) doExport(true);
    return;
  }
  if (!r.ok) { showResult(`<span class="ng">${esc(j.error || r.statusText)}</span>`); return; }
  const ok = j.results.filter((x) => x.ok).length;
  let html = `<b>${ok} / ${j.results.length} 個を保存しました</b> — <a id="openFolder">${esc(j.folder)}</a>\n`;
  html += j.results.map((x) => x.ok
    ? `<span class="ok">✓ ${esc(x.file)}（${x.size[0]}×${x.size[1]} px）</span>`
    : `<span class="ng">✗ ${esc(x.file)}: ${esc(x.error)}</span>`).join("\n");
  if (j.warning) html += `\n<span class="ng">${esc(j.warning)}</span>`;
  showResult(html);
  $("openFolder").onclick = () => fetch("/api/open-folder", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ folder: j.folder }),
  });
}

function showResult(html) { $("result").innerHTML = html; }
function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

// ------------------------------------------------------------------ ファイル・キー操作

$("fileInput").onchange = (e) => { loadFile(e.target.files[0]); e.target.value = ""; };

["dragenter", "dragover"].forEach((t) => document.addEventListener(t, (e) => {
  e.preventDefault(); $("drop").classList.add("over");
}));
["dragleave", "drop"].forEach((t) => document.addEventListener(t, (e) => {
  e.preventDefault(); $("drop").classList.remove("over");
}));
document.addEventListener("drop", (e) => {
  const f = e.dataTransfer?.files?.[0];
  if (f) loadFile(f);
});

document.addEventListener("keydown", (e) => {
  if (!data) return;
  const typing = ["INPUT", "SELECT", "TEXTAREA"].includes(document.activeElement?.tagName);
  if (typing) return;
  const mod = e.metaKey || e.ctrlKey;
  if (e.key === "Escape") { if (currentKey) setCurrent(null); }
  else if (mod && e.key.toLowerCase() === "a") { e.preventDefault(); addMany(data.slides[cur].shapes); }
  else if (mod && e.key.toLowerCase() === "z") {
    e.preventDefault();
    if (history.length) { const h = history.pop(); selected = h.list; currentKey = h.current; drawOverlay(); renderList(); }
  } else if (["ArrowDown", "ArrowRight", "PageDown"].includes(e.key)) { e.preventDefault(); showSlide(cur + 1); }
  else if (["ArrowUp", "ArrowLeft", "PageUp"].includes(e.key)) { e.preventDefault(); showSlide(cur - 1); }
});

window.addEventListener("resize", fitStage);
$("slideImg").addEventListener("load", drawOverlay);
init();
