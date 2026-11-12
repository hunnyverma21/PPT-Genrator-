const $ = id => document.getElementById(id);
let slides = [{
  title: "Welcome to my presentation",
  body: "Introduce your topic\nShare the main objective\nExplain what your audience will learn",
  layout: "content"
}];
let active = 0;

const themes = {
  midnight: { bg: "142746", accent: "8FB5F5", text: "FFFFFF", muted: "DCE7F7" },
  violet: { bg: "30204F", accent: "D0B7FF", text: "FFFFFF", muted: "EEE5FF" },
  forest: { bg: "123C37", accent: "9CE4D4", text: "FFFFFF", muted: "D8F5EC" },
  minimal: { bg: "F8FAFC", accent: "5B4BC4", text: "172033", muted: "475569" }
};

function currentSlide() { return slides[active]; }
function setStatus(text, error = false) {
  $("status").textContent = text;
  $("status").style.color = error ? "#fda4af" : "#8f9db2";
}
function syncFromEditor() {
  const s = currentSlide();
  s.title = $("slideTitle").value;
  s.body = $("slideBody").value;
  s.layout = $("layout").value;
}
function render() {
  const s = currentSlide();
  $("editorTitle").textContent = `Edit slide ${active + 1}`;
  $("slideTitle").value = s.title;
  $("slideBody").value = s.body;
  $("layout").value = s.layout || "content";
  $("slideCount").textContent = `${slides.length} slide${slides.length === 1 ? "" : "s"}`;
  $("deleteSlide").disabled = slides.length === 1;
  $("deleteSlide").style.opacity = slides.length === 1 ? ".4" : "1";

  $("slideList").innerHTML = "";
  slides.forEach((slide, i) => {
    const button = document.createElement("button");
    button.className = "slide-item" + (i === active ? " active" : "");
    const thumb = document.createElement("span");
    thumb.className = "thumb";
    thumb.textContent = `0${i + 1}`;
    const info = document.createElement("span");
    const title = document.createElement("strong");
    title.textContent = slide.title || "Untitled slide";
    const sub = document.createElement("small");
    sub.textContent = slide.layout === "title" ? "Title slide" : "Title + bullets";
    info.append(title, sub);
    button.append(thumb, info);
    button.addEventListener("click", () => { syncFromEditor(); active = i; render(); });
    $("slideList").append(button);
  });

  const theme = $("theme").value;
  $("preview").className = `slide-preview theme-${theme}`;
  $("previewTitle").textContent = s.title || "Untitled slide";
  $("previewBullets").innerHTML = "";
  const points = s.body.split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  points.forEach(point => {
    const li = document.createElement("li");
    li.textContent = point;
    $("previewBullets").append(li);
  });
  $("previewBullets").style.display = s.layout === "title" ? "none" : "block";
  $("preview").querySelector(".preview-kicker").textContent = $("presentationTitle").value.trim() || "MY PRESENTATION";
  $("preview").querySelector(".preview-number").textContent = String(active + 1).padStart(2, "0");
}
function addSlide() {
  syncFromEditor();
  slides.push({ title: "New slide", body: "Add your first point\nAdd another point", layout: "content" });
  active = slides.length - 1;
  render();
  $("slideTitle").focus();
}
function addPptText(slide, text, opts) {
  if (text) slide.addText(text, opts);
}
async function exportPptx() {
  syncFromEditor();
  if (typeof pptxgen === "undefined") {
    setStatus("PowerPoint library did not load. Check your internet connection and refresh.", true);
    return;
  }
  const title = $("presentationTitle").value.trim() || "My Presentation";
  const theme = themes[$("theme").value];
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "SlideCraft";
  pptx.subject = "Presentation created with SlideCraft";
  pptx.title = title;
  pptx.company = "";
  pptx.lang = "en-US";

  slides.forEach((item, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: theme.bg };
    const isTitle = item.layout === "title";
    const heading = item.title || `Slide ${index + 1}`;
    const bodyLines = item.body.split(/\r?\n/).map(x => x.trim()).filter(Boolean);
    addPptText(slide, title.toUpperCase(), {
      x: 0.65, y: 0.42, w: 11.8, h: 0.25, fontFace: "Aptos",
      fontSize: 9, bold: true, charSpacing: 2, color: theme.accent, margin: 0
    });
    addPptText(slide, heading, {
      x: isTitle ? 0.9 : 0.8, y: isTitle ? 2.0 : 1.15, w: 11.3, h: isTitle ? 1.25 : 1.0,
      fontFace: "Aptos Display", fontSize: isTitle ? 34 : 29, bold: true,
      color: theme.text, breakLine: false, valign: "mid", margin: 0.04,
      fit: "shrink"
    });
    if (!isTitle && bodyLines.length) {
      slide.addText(bodyLines.map(line => ({ text: line, options: { bullet: { indent: 18 }, hanging: 4, breakLine: true } })), {
        x: 1.0, y: 2.45, w: 10.8, h: 3.65, fontFace: "Aptos",
        fontSize: 19, color: theme.muted, paraSpaceAfterPt: 15,
        breakLine: false, valign: "top", margin: 0.08, fit: "shrink"
      });
    }
    slide.addText(String(index + 1).padStart(2, "0"), {
      x: 12.1, y: 6.85, w: 0.55, h: 0.25, fontFace: "Aptos",
      fontSize: 10, color: theme.accent, align: "right", margin: 0
    });
  });

  try {
    setStatus("Generating your PowerPoint file…");
    await pptx.writeFile({ fileName: `${title.replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "_") || "presentation"}.pptx` });
    setStatus("PowerPoint downloaded successfully.");
  } catch (error) {
    setStatus("Could not create the PowerPoint. Please try again.", true);
  }
}

$("addSlide").addEventListener("click", addSlide);
$("addAnother").addEventListener("click", addSlide);
$("deleteSlide").addEventListener("click", () => {
  if (slides.length <= 1) return;
  slides.splice(active, 1);
  active = Math.min(active, slides.length - 1);
  render();
});
["slideTitle", "slideBody", "presentationTitle", "theme", "layout"].forEach(id => {
  $(id).addEventListener("input", () => { syncFromEditor(); render(); });
  $(id).addEventListener("change", () => { syncFromEditor(); render(); });
});
$("download").addEventListener("click", exportPptx);
$("downloadTop").addEventListener("click", exportPptx);
render();
