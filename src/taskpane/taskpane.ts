import { getTableSnapshot, applyFilterPlan, clearAllFilters } from "./excel";

import { parseAndApply } from "./parser";

/* global Office */

Office.onReady(() => {
  const applyButton = document.getElementById("applyBtn");
  const clearButton = document.getElementById("clearBtn");
  const status = document.getElementById("status");

  applyButton?.addEventListener("click", async () => {
    const prompt = (document.getElementById("promptInput") as HTMLTextAreaElement).value;

    if (status) status.innerText = "Working...";

    try {
      const { headers, rows } = await getTableSnapshot();
      const { plan, notes } = parseAndApply(headers, rows, prompt);

      // visibleCount comes straight from Excel after the filter is applied,
      // so it always matches exactly what's showing in the sheet.
      const visibleCount = await applyFilterPlan(plan);

      if (status) {
        status.innerText = `${visibleCount} wine${visibleCount === 1 ? "" : "s"} found — ${notes.join(" • ")}`;
      }
    } catch (err: any) {
      if (status) status.innerText = `Error: ${err?.message || err}`;
    }
  });

  clearButton?.addEventListener("click", async () => {
    try {
      await clearAllFilters();
      if (status) status.innerText = "Filters cleared — showing all wines";
    } catch (err: any) {
      if (status) status.innerText = `Error: ${err?.message || err}`;
    }
  });
});