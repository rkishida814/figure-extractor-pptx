"use strict";

const $ = (id) => document.getElementById(id);
const EXTS = ["png", "tif", "bmp", "jpg"];
const LS = {
  get(k, d) { try { const v = localStorage.getItem("fx." + k); return v === null ? d : v; } catch { return d; } },
  set(k, v) { try { localStorage.setItem("fx." + k, v); } catch { /* なし */ } },
};

let data = null;          // /api/upload の結果
let cur = 0;              // 表示中のスライド
let selected = [];        // [{key, slide, id}] 選んだ順
let custom = {};          // key -> {name?, ext?}（オブジェクトごとに覚えておく）
let history = [];         // 選択の取り消し用
let defaultExt = LS.get("ext", "png");
let hoverKey = null;

const keyOf = (slide, id) => `${slide}:${id}`;
const shapeOf = (slide, id) => data.slides[slide].shapes.find((s) => s.id === id);

// ------------------------------------------------------------------ 読み込み

async function init() {
  $("folder").value = LS.get("folder", "");
  $("dpi").value = LS.get("dpi", "300");
  $("transparent").checked = LS.get("transparent", "0") === "1";
  try {
    const st = await (await fetch("/api/status")).json();
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
    cur = 0; selected = []; custom = {}; history = [];
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
  const order = new Map(selected.map((it, k) => [it.key, k + 1]));
  for (const sh of data.slides[cur].shapes) {
    const key = keyOf(cur, sh.id);
    const n = order.get(key);
    if (!n && key !== hoverKey) continue;
    const b = document.createElement("div");
    b.className = "box " + (n ? "sel" : "hover");
    Object.assign(b.style, {
      left: sh.x * s + "px", top: sh.y * s + "px",
      width: Math.max(2, sh.w * s) + "px", height: Math.max(2, sh.h * s) + "px",
      transform: sh.rot ? `rotate(${sh.rot}deg)` : "",
    });
    if (n) { const t = document.createElement("span"); t.className = "num"; t.textContent = n; b.appendChild(t); }
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

function pushHistory() {
  history.push(selected.map((it) => ({ ...it })));
  if (history.length > 100) history.shift();
}

function setSelection(list) {
  pushHistory();
  selected = list;
  drawOverlay();
  renderList();
}

function item(slide, id) { return { key: keyOf(slide, id), slide, id }; }

function toggle(slide, id) {
  const key = keyOf(slide, id);
  if (selected.some((it) => it.key === key)) setSelection(selected.filter((it) => it.key !== key));
  else setSelection([...selected, item(slide, id)]);
}

function addMany(shapes) {
  const have = new Set(selected.map((it) => it.key));
  setSelection([...selected, ...shapes.filter((sh) => !have.has(keyOf(cur, sh.id))).map((sh) => item(cur, sh.id))]);
}

let drag = null;

$("stage").addEventListener("mousedown", (ev) => {
  if (ev.button !== 0 || !data) return;
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
    const p = toEmu(ev);
    const x0 = Math.min(p.x, d.start.x), x1 = Math.max(p.x, d.start.x);
    const y0 = Math.min(p.y, d.start.y), y1 = Math.max(p.y, d.start.y);
    const inside = data.slides[cur].shapes.filter((sh) => {
      const b = aabb(sh);
      return b.x >= x0 && b.y >= y0 && b.x + b.w <= x1 && b.y + b.h <= y1;
    });
    if (d.add) addMany(inside);
    else setSelection(inside.map((sh) => item(cur, sh.id)));
    return;
  }
  if (d.add && !d.alt) {   // ⌘ 等だけ：最前面を追加・解除
    const sh = hit(d.start);
    lastPick = sh ? { slide: cur, x: d.start.x, y: d.start.y, id: sh.id } : null;
    if (sh) toggle(cur, sh.id);
    return;
  }
  const sh = pickAt(d.start, d.alt);
  if (d.add) {             // ⌘+⌥：下のオブジェクトを追加
    if (sh && !selected.some((it) => it.key === keyOf(cur, sh.id))) setSelection([...selected, item(cur, sh.id)]);
    return;
  }
  if (sh) {
    const key = keyOf(cur, sh.id);
    if (selected.length === 1 && selected[0].key === key) return;
    setSelection([item(cur, sh.id)]);
  } else if (selected.length) {
    setSelection([]);
  }
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
    const sh = shapeOf(it.slide, it.id);
    const li = document.createElement("li");
    li.dataset.key = it.key;

    const idx = document.createElement("span");
    idx.className = "idx"; idx.textContent = k + 1;

    const crop = document.createElement("div");
    crop.className = "crop";
    const b = aabb(sh), W = 96, H = 64;
    const z = Math.min(W / Math.max(b.w, 1), H / Math.max(b.h, 1));
    Object.assign(crop.style, {
      backgroundImage: `url(${data.slides[it.slide].image})`,
      backgroundSize: `${data.slideW * z}px ${data.slideH * z}px`,
      backgroundPosition: `${-b.x * z + (W - b.w * z) / 2}px ${-b.y * z + (H - b.h * z) / 2}px`,
    });
    crop.title = "このスライドを表示";
    crop.onclick = () => showSlide(it.slide);

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
      if (e.key === "Enter" || (e.key === "ArrowDown")) { e.preventDefault(); ol.children[k + 1]?.querySelector("input")?.focus(); }
      if (e.key === "ArrowUp") { e.preventDefault(); ol.children[k - 1]?.querySelector("input")?.focus(); }
    });
    const ext = document.createElement("select");
    ext.className = "ext";
    EXTS.forEach((e) => { const o = document.createElement("option"); o.value = e; o.textContent = "." + e; ext.appendChild(o); });
    ext.value = extOf(it);
    ext.onchange = () => { custom[it.key] = custom[it.key] || {}; custom[it.key].ext = ext.value; };
    row.append(name, ext);
    const src = document.createElement("div");
    src.className = "src";
    src.textContent = `スライド${it.slide + 1}・${sh.kind}・${sh.name}`;
    fields.append(row, src);

    const rm = document.createElement("button");
    rm.className = "rm"; rm.textContent = "×"; rm.title = "一覧から外す";
    rm.onclick = () => setSelection(selected.filter((x) => x.key !== it.key));

    li.append(idx, crop, fields, rm);
    li.addEventListener("mouseenter", () => { if (it.slide === cur) { hoverKey = it.key; drawOverlay(); } });
    li.addEventListener("mouseleave", () => { hoverKey = null; drawOverlay(); });
    ol.appendChild(li);
  });
}

function highlightRow(key) {
  [...$("list").children].forEach((li) => li.classList.toggle("hl", li.dataset.key === key));
}

$("bulkExt").onchange = (e) => {
  const v = e.target.value;
  if (!v) return;
  defaultExt = v;
  LS.set("ext", v);
  selected.forEach((it) => { custom[it.key] = custom[it.key] || {}; custom[it.key].ext = v; });
  e.target.value = "";
  renderList();
};

$("clearBtn").onclick = () => setSelection([]);
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
  if (j.folder) { $("folder").value = j.folder; LS.set("folder", j.folder); }
};
$("folder").onchange = () => LS.set("folder", $("folder").value);
$("dpi").onchange = () => LS.set("dpi", $("dpi").value);
$("transparent").onchange = () => LS.set("transparent", $("transparent").checked ? "1" : "0");

$("saveBtn").onclick = () => doExport(false);

async function doExport(overwrite) {
  if (!data || !selected.length) return;
  const folder = $("folder").value.trim() || $("folder").placeholder;
  $("folder").value = folder;
  LS.set("folder", folder);
  const items = selected.map((it, k) => ({ slide: it.slide, shapeId: it.id, name: nameOf(it, k).trim(), ext: extOf(it) }));
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
  if (e.key === "Escape") { if (selected.length) setSelection([]); }
  else if (mod && e.key.toLowerCase() === "a") { e.preventDefault(); addMany(data.slides[cur].shapes); }
  else if (mod && e.key.toLowerCase() === "z") {
    e.preventDefault();
    if (history.length) { selected = history.pop(); drawOverlay(); renderList(); }
  } else if (["ArrowDown", "ArrowRight", "PageDown"].includes(e.key)) { e.preventDefault(); showSlide(cur + 1); }
  else if (["ArrowUp", "ArrowLeft", "PageUp"].includes(e.key)) { e.preventDefault(); showSlide(cur - 1); }
});

window.addEventListener("resize", fitStage);
$("slideImg").addEventListener("load", drawOverlay);
init();
