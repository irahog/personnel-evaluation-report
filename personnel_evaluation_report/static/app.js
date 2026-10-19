"use strict";
document.addEventListener("DOMContentLoaded", () => {
  // Confirmation for destructive / state-clearing actions.
  document.querySelectorAll("form[data-confirm]").forEach((f) => {
    f.addEventListener("submit", (e) => { if (!window.confirm(f.dataset.confirm)) e.preventDefault(); });
  });
  // Prevent double submits; show a busy note where useful.
  document.querySelectorAll("form").forEach((f) => {
    f.addEventListener("submit", (e) => {
      if (e.defaultPrevented) return;
      f.querySelectorAll("button").forEach((b) => { b.disabled = true; });
      if (f.dataset.busy) { const p = document.createElement("p"); p.className = "hint"; p.textContent = f.dataset.busy; f.append(p); }
    });
  });
  // Photo: show the chosen picture immediately (it is uploaded with the form).
  const photoInput = document.getElementById("photo-input");
  if (photoInput) {
    photoInput.addEventListener("change", () => {
      const f = photoInput.files && photoInput.files[0];
      const img = document.getElementById("photo-img");
      const empty = document.getElementById("photo-empty");
      if (!f || !img) return;
      img.src = URL.createObjectURL(f);
      img.hidden = false;
      if (empty) empty.hidden = true;
    });
  }
  // Table appearance: live preview in the form (the server renders the same values for view/print/Excel).
  // The label column of «اطلاعات فردی/پرسنلی» belongs to neither part (only the title bar is header).
  const PART_TARGETS = { header: "h3, thead th", body: "tbody td, table.kv td, .note" };
  function applyPart(section, fs) {
    const v = (p) => fs.querySelector(`[data-prop="${p}"]`);
    const scale = parseFloat(v("size").selectedOptions[0].dataset.scale);
    section.querySelectorAll(PART_TARGETS[fs.dataset.part]).forEach((el) => {
      if (el.closest(".style-panel")) return;
      el.style.color = v("fg").value;
      el.style.background = v("bg").value;
      el.style.fontWeight = v("bold").checked ? "700" : "400";
      el.style.fontStyle = v("italic").checked ? "italic" : "normal";
      el.style.textDecoration = v("underline").checked ? "underline" : "none";
      el.style.fontSize = scale === 1 ? "" : (el.tagName === "H3" ? `calc(1.05rem * ${scale})` : `${scale}em`);
    });
  }
  document.querySelectorAll(".style-panel").forEach((panel) => {
    const section = panel.closest("[data-sec]");
    panel.querySelectorAll("fieldset").forEach((fs) => {
      fs.addEventListener("input", () => applyPart(section, fs));
      fs.addEventListener("change", () => applyPart(section, fs));
    });
    panel.querySelector(".style-reset").addEventListener("click", () => {
      panel.querySelectorAll("[data-prop]").forEach((el) => {
        if (el.type === "checkbox") el.checked = el.dataset.default === "1";
        else el.value = el.dataset.default;
      });
      panel.querySelectorAll("fieldset").forEach((fs) => applyPart(section, fs));
      panel.closest("form").dispatchEvent(new Event("input"));
    });
  });
  // Print.
  document.querySelectorAll("[data-print]").forEach((b) => b.addEventListener("click", () => window.print()));
  // Warn before leaving the form with unsaved edits (nothing is saved automatically).
  const guarded = document.querySelector("form[data-dirty-guard]");
  if (guarded) {
    let dirty = false;
    guarded.addEventListener("input", () => { dirty = true; });
    guarded.addEventListener("submit", () => { dirty = false; });
    window.addEventListener("beforeunload", (e) => { if (dirty) { e.preventDefault(); e.returnValue = ""; } });
  }
  // Font availability (only Times New Roman is not bundled).
  const probe = document.getElementById("font-probe");
  if (probe && document.fonts) {
    document.fonts.ready.then(() => {
      if (!document.fonts.check(`16px "${probe.dataset.font}"`, "آزمون")) {
        document.getElementById("font-warning").hidden = false;
      }
    });
  }
});
