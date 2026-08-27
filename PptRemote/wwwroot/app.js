const notesEl = document.getElementById("notes");
const previewEl = document.getElementById("preview");
const previewImg = document.getElementById("previewImg");
const previewCap = document.getElementById("previewCap");
const counter = document.getElementById("counter");
const titleEl = document.getElementById("title");
const pill = document.getElementById("pill");
const dock = document.getElementById("dock");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");
const tabNotes = document.getElementById("tabNotes");
const tabNext = document.getElementById("tabNext");
const gridBtn = document.getElementById("gridBtn");
const gridClose = document.getElementById("gridClose");
const grid = document.getElementById("grid");
const gridList = document.getElementById("gridList");
const more = document.getElementById("more");
const moreBtn = document.getElementById("moreBtn");
const moreClose = document.getElementById("moreClose");
const offline = document.getElementById("offline");
const offlineHint = document.getElementById("offlineHint");
const retryBtn = document.getElementById("retryBtn");
const blackBtn = document.getElementById("blackBtn");
const endTopBtn = document.getElementById("endTopBtn");

let state = null;
let mode = localStorage.getItem("ppt-mode") === "next" ? "next" : "notes";
let lastSig = "";
let lastGridStruct = "";
let lastGridVis = "";
let misses = 0;
let touchX = 0;
let touchY = 0;
let swiping = false;

offlineHint.textContent = "Connected as " + location.host;

function cleanNotes(raw) {
  return String(raw || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function showOffline(on) {
  offline.hidden = !on;
  offline.style.setProperty("display", on ? "flex" : "none", "important");
}

function thumbUrl(index) {
  const v = state && state.thumbsVersion ? state.thumbsVersion : 0;
  return "/thumbs/" + index + ".png?v=" + v;
}

function pull() {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "/live.js?t=" + Date.now();
    s.onload = () => {
      s.remove();
      if (window.__BOOT) resolve(window.__BOOT);
      else reject();
    };
    s.onerror = () => {
      s.remove();
      reject();
    };
    document.head.appendChild(s);
  });
}

function go(name, extra) {
  const img = new Image();
  img.src = "/go/" + name + "?t=" + Date.now() + (extra || "");
}

async function poll() {
  try {
    const next = await pull();
    misses = 0;
    apply(next);
    showOffline(false);
    if (!grid.hidden) paintGrid();
  } catch {
    misses += 1;
    if (!state && misses >= 8) showOffline(true);
  }
}

function send(name, extra) {
  go(name, extra);
  setTimeout(poll, 180);
}

function setMode(next) {
  mode = next;
  localStorage.setItem("ppt-mode", next);
  tabNotes.setAttribute("aria-selected", next === "notes" ? "true" : "false");
  tabNext.setAttribute("aria-selected", next === "next" ? "true" : "false");
  notesEl.hidden = next !== "notes";
  previewEl.hidden = next !== "next";
  document.getElementById("app").classList.toggle("next-mode", next === "next");
  paintPreview();
}

function paintPreview() {
  if (!state || mode !== "next") return;
  if (state.aheadReady) {
    previewImg.onerror = () => {
      const n = state.nextIndex;
      if (n) previewImg.src = thumbUrl(n);
    };
    previewImg.src = "/ahead.png?v=" + (state.aheadVersion || 0);
    previewCap.textContent = state.aheadLabel || "Next click";
    return;
  }
  const n = state.nextIndex;
  if (!n) {
    previewImg.removeAttribute("src");
    previewCap.textContent = "Last slide";
    return;
  }
  previewImg.src = thumbUrl(n);
  previewCap.textContent = "Next slide · " + n + " / " + state.total;
}

function paintNotes() {
  if (!state) return;
  const text = cleanNotes(state.notes);
  if (!state.connected) {
    notesEl.textContent = state.message || "Open a deck on the PC, then start the slideshow.";
    notesEl.classList.add("empty");
    return;
  }
  if (!text) {
    notesEl.textContent = "No notes on this slide.";
    notesEl.classList.add("empty");
    return;
  }
  notesEl.classList.remove("empty");
  notesEl.textContent = text;
}

function paintDock() {
  const live = !!(state && state.slideshow);
  endTopBtn.hidden = false;
  endTopBtn.disabled = !live;
  endTopBtn.style.opacity = live ? "1" : "0.35";
  if (!state || !state.connected) {
    dock.classList.add("single");
    prevBtn.hidden = true;
    nextBtn.hidden = false;
    nextBtn.textContent = "Waiting…";
    nextBtn.disabled = true;
    nextBtn.className = "nav ghost";
    return;
  }
  if (!state.slideshow) {
    dock.classList.add("single");
    prevBtn.hidden = true;
    nextBtn.hidden = false;
    nextBtn.disabled = false;
    nextBtn.textContent = "Start slideshow";
    nextBtn.className = "nav solid";
    return;
  }
  dock.classList.remove("single");
  prevBtn.hidden = false;
  nextBtn.hidden = false;
  prevBtn.disabled = false;
  nextBtn.disabled = false;
  prevBtn.textContent = "Prev";
  nextBtn.textContent = "Next";
  prevBtn.className = "nav ghost";
  nextBtn.className = "nav solid";
}

function apply(next) {
  const sig = [
    next.connected,
    next.slideshow,
    next.black,
    next.index,
    next.total,
    next.notes,
    next.nextIndex,
    next.thumbsVersion,
    next.thumbsReady,
    next.clickIndex,
    next.clickCount,
    next.aheadVersion,
    next.aheadReady,
    next.aheadLabel,
    next.message,
    next.title,
  ].join("|");
  if (sig === lastSig) return;
  const notesChanged = !state || state.notes !== next.notes || state.index !== next.index || state.connected !== next.connected;
  state = next;
  lastSig = sig;
  counter.textContent = next.total ? next.index + " / " + next.total : "– / –";
  titleEl.textContent = next.title || next.message || "PPT Remote";
  pill.hidden = !next.slideshow;
  pill.textContent = next.black ? "BLACK" : next.white ? "WHITE" : "LIVE";
  blackBtn.textContent = next.black ? "Resume from black" : "Black screen";
  if (notesChanged) paintNotes();
  paintPreview();
  paintDock();
  showOffline(false);
}

function paintGrid() {
  if (!state) return;
  const slides = state.slides || [];
  const structSig = state.total + "|" + JSON.stringify(slides);
  if (structSig !== lastGridStruct) {
    lastGridStruct = structSig;
    lastGridVis = "";
    gridList.innerHTML = "";
    for (const slide of slides) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "card";
      btn.dataset.index = String(slide.index);
      if (slide.hidden) btn.classList.add("hidden-slide");
      const img = document.createElement("img");
      img.alt = "Slide " + slide.index;
      img.decoding = "async";
      img.addEventListener("error", () => {
        img.style.opacity = "0";
      });
      img.addEventListener("load", () => {
        img.style.opacity = "1";
      });
      const cap = document.createElement("span");
      cap.textContent = slide.hidden ? slide.index + "  hidden" : String(slide.index);
      btn.append(img, cap);
      btn.addEventListener("click", () => jump(slide.index));
      gridList.append(btn);
    }
  }
  const visSig = state.thumbsVersion + "|" + state.thumbsReady + "|" + state.index;
  if (visSig === lastGridVis) return;
  lastGridVis = visSig;
  for (const btn of gridList.children) {
    const i = +btn.dataset.index;
    btn.classList.toggle("current", i === state.index);
    const img = btn.querySelector("img");
    if (!img) continue;
    const url = thumbUrl(i);
    if (img.getAttribute("src") !== url) img.src = url;
  }
}

function jump(index) {
  grid.hidden = true;
  send("goto", "&n=" + index);
}

function endShow() {
  more.hidden = true;
  send("end");
}

tabNotes.addEventListener("click", () => setMode("notes"));
tabNext.addEventListener("click", () => setMode("next"));
gridBtn.addEventListener("click", () => {
  more.hidden = true;
  grid.hidden = false;
  lastGridStruct = "";
  lastGridVis = "";
  paintGrid();
});
gridClose.addEventListener("click", () => {
  grid.hidden = true;
});
moreBtn.addEventListener("click", () => {
  grid.hidden = true;
  more.hidden = false;
});
moreClose.addEventListener("click", () => {
  more.hidden = true;
});
retryBtn.addEventListener("click", () => {
  showOffline(false);
  poll();
});
prevBtn.addEventListener("click", () => send("prev"));
nextBtn.addEventListener("click", () => {
  if (state && state.connected && !state.slideshow) {
    send("start");
    return;
  }
  send("next");
});
endTopBtn.addEventListener("click", endShow);
document.getElementById("endBtn").addEventListener("click", endShow);
blackBtn.addEventListener("click", () => send("black"));
document.getElementById("whiteBtn").addEventListener("click", () => send("white"));
document.getElementById("firstBtn").addEventListener("click", () => {
  more.hidden = true;
  send("first");
});
document.getElementById("lastBtn").addEventListener("click", () => {
  more.hidden = true;
  send("last");
});

previewEl.addEventListener("touchstart", (e) => {
  const t = e.changedTouches[0];
  touchX = t.clientX;
  touchY = t.clientY;
  swiping = true;
}, { passive: true });

previewEl.addEventListener("touchend", (e) => {
  if (!swiping) return;
  swiping = false;
  if (mode !== "next") return;
  if (!state || !state.slideshow) return;
  const t = e.changedTouches[0];
  const dx = t.clientX - touchX;
  const dy = t.clientY - touchY;
  if (Math.abs(dx) < 72) return;
  if (Math.abs(dx) < Math.abs(dy) * 1.2) return;
  if (dx < 0) send("next");
  else send("prev");
}, { passive: true });

previewEl.addEventListener("click", (e) => {
  e.preventDefault();
  e.stopPropagation();
}, true);

document.addEventListener("gesturestart", (e) => e.preventDefault());
document.addEventListener("dblclick", (e) => e.preventDefault());

showOffline(false);
setMode(mode);
const view = new URLSearchParams(location.search).get("view");
if (view === "next") setMode("next");
if (window.__BOOT) apply(window.__BOOT);
if (view === "grid") {
  grid.hidden = false;
  lastGridStruct = "";
  lastGridVis = "";
  paintGrid();
}
poll();
setInterval(poll, 700);

if (navigator.wakeLock) {
  const lock = () => navigator.wakeLock.request("screen").catch(() => {});
  lock();
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") lock();
  });
}
