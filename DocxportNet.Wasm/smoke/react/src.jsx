import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { createDocxport } from "docxport";

function App() {
  const [state, setState] = useState({ status: "loading", output: "" });
  useEffect(() => {
    let active = true;
    (async () => {
      const base = new URL(`${import.meta.env.BASE_URL}docxport/`, window.location.origin);
      const input = new Uint8Array(await (await fetch(`${import.meta.env.BASE_URL}sample.docx`)).arrayBuffer());
      const api = await createDocxport({ assetBaseUrl: base });
      const output = await api.export(input, { format: "text", text: { trackedChangeMode: "accept" } });
      if (active) setState({ status: "ready", output });
    })().catch(error => active && setState({ status: "error", output: String(error) }));
    return () => { active = false; };
  }, []);
  return <main data-status={state.status}><pre>{state.output}</pre></main>;
}

createRoot(document.getElementById("root")).render(<App />);
