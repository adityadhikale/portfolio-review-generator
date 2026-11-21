# CREST Portfolio Review: website front end

Vue 3 + Vite + TypeScript, with shadcn-vue components (`src/components/ui`), Motion for Vue animation and Tailwind.
The colours and fonts in `src/style.css` come from the CREST review deck.

- `npm install` then `npm run build`: writes `dist/`, which `python serve.py` (in the folder above) serves.
- `npm run dev`: live-reload development at http://localhost:5173; `/api` is forwarded to `python serve.py` on :8000.
- Add a shadcn-vue component: `npx shadcn-vue@latest add <name>`.

Screens: `src/views/NewRun.vue` (choose a step, upload), `src/views/RunView.vue` (queue, live progress, results),
`src/views/History.vue`. The API client is `src/api.ts`; the API itself is `crest/api.py`.
