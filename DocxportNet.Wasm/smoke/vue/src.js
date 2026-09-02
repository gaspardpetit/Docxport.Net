import { createApp, h, ref } from "vue";
import { createDocxport } from "docxport";

createApp({
  setup() {
    const status = ref("loading");
    const output = ref("");
    (async () => {
      const base = new URL(`${import.meta.env.BASE_URL}docxport/`, window.location.origin);
      const input = new Uint8Array(await (await fetch(`${import.meta.env.BASE_URL}sample.docx`)).arrayBuffer());
      const api = await createDocxport({ assetBaseUrl: base });
      let progressReports = 0;
      output.value = await api.export(input, {
        format: "html",
        preset: "rich",
        onProgress: () => progressReports++
      });
      status.value = progressReports > 0 ? "ready" : "missing-progress";
    })().catch(error => { status.value = "error"; output.value = String(error); });
    return () => h("main", { "data-status": status.value }, [h("pre", output.value)]);
  }
}).mount("#app");
