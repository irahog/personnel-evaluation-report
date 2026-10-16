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
