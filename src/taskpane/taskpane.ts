import { runRedaction } from "../redactor/redactor";

function setStatus(text: string, kind: "idle" | "success" | "error" = "idle") {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = text;
  el.style.color =
    kind === "success" ? "var(--success)" :
    kind === "error" ? "var(--danger)" :
    "var(--muted)";
}

function showSummary(result: { emails:number; phones:number; ssns:number; cards:number; note?:string }) {
  const summary = document.getElementById("summary");
  if (!summary) return;
  summary.classList.remove("hidden");

  (document.getElementById("emails") as HTMLElement).textContent = String(result.emails);
  (document.getElementById("phones") as HTMLElement).textContent = String(result.phones);
  (document.getElementById("ssns") as HTMLElement).textContent = String(result.ssns);
  (document.getElementById("cards") as HTMLElement).textContent = String(result.cards);

  const noteEl = document.getElementById("note");
  if (noteEl) noteEl.textContent = result.note ?? "";
}

Office.onReady((info: any) => {
  if (info.host !== Office.HostType.Word) return;

  const btn = document.getElementById("redactBtn") as HTMLButtonElement | null;
  if (!btn) return;

  btn.addEventListener("click", async () => {
    btn.disabled = true;
    setStatus("Processing…");

    try {
      const result = await runRedaction();
      setStatus("Redaction complete.", "success");
      showSummary(result);
    } catch (e: any) {
      console.error(e);
      setStatus(`Error: ${e?.message ?? String(e)}`, "error");
    } finally {
      btn.disabled = false;
    }
  });
});
