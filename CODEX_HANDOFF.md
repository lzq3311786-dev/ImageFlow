# ImageFlow Codex Handoff

## Current Snapshot
- Project: ImageFlow / `img-studio`
- Current release: `2.1.6`
- Latest source commit: `b75f476 Release 2.1.6`
- Repository: `https://github.com/lzq3311786-dev/ImageFlow.git`
- Main runtime: Electron CommonJS app.
- Main files:
  - `main.js`: Electron main process, IPC handlers, file IO, template rendering orchestration, product publish export, update flow, print AI workflow.
  - `图片工作台.html`: single-page UI, styles, renderer logic, all module UI code.
  - `template_renderer.py`: Python template renderer used by smart template generation.
  - `template_mask_points.py`: helper for template mask coordinate logic.
  - `templates/`: bundled template assets.
  - `watermarks/`: bundled watermark presets/assets.
  - `python-runtime/`: bundled Python runtime for packaged app.

## Setup
1. Install Node.js on Windows.
2. From repo root:
   ```powershell
   npm install
   npm run start
   ```
3. Build installer:
   ```powershell
   npm run build
   ```
4. Publish GitHub release:
   ```powershell
   $env:GH_TOKEN = "<github token>"
   npm run release:github
   ```

## Packaging And Update Flow
- `package.json` version is the source of release version.
- `npm run release:github` runs `electron-builder --publish always`.
- Installer output is under `dist/`, for example `dist/ImageFlow Setup 2.1.6.exe`.
- GitHub auto update uses `electron-updater` and the `build.publish` config in `package.json`.
- For a new release, update both `package.json` and `package-lock.json`, then build and publish.
- Do not publish the same version twice unless intentionally replacing release assets.

## Main Modules
- Compress export: config stored in `compress-config.json`.
- One-click classify: config stored in `classify-config.json`; product prefix rules are managed in renderer settings.
- Smart slice:
  - UI and algorithms live mostly in `图片工作台.html` around the `sliceState` section.
  - Results are saved through `slice:save-results` in `main.js`.
  - Current slice workflow supports task selection, manual/auto cuts, shrink-edge preview, result viewer, import to template.
- Smart template:
  - Template config IPC prefix: `template:*`.
  - Template root defaults to external app directory in dev and userData in packaged mode.
  - Uses Python renderer through `template_renderer.py`.
  - Output can be imported into product publish.
- Product publish:
  - Config: `product-publish-config.json`.
  - Data: `product-publish-data.json`.
  - Supports AI title generation, prompt presets, AI/OSS presets, product type mapping, Temu XLSX export.
  - OSS upload can convert local images to URLs before export.
- Print AI / 印花裂变:
  - Config: `print-ai-config.json`.
  - Data: `print-ai-data.json`.
  - Storage root: `userData/print-ai`.
  - IPC prefix: `print-ai:*`.
  - Current workflow: import product image -> extract clean print image -> generate same-series 3x3 variation image -> import variation result into smart slice.
  - Model detection for print AI uses the real `/models` response only; it does not append product-publish fallback model lists.

## User Data Locations
In development this app currently writes under:
```text
C:\Users\Administrator\AppData\Roaming\img-studio
```
Important files:
- `compress-config.json`
- `classify-config.json`
- `slice-config.json`
- `template-config.json`
- `product-publish-config.json`
- `product-publish-data.json`
- `print-ai-config.json`
- `print-ai-data.json`
- `print-ai/`

API keys, OSS keys, user tasks, and generated images are local userData state. Do not assume they exist in the repository.

## Development Rules For Future Codex
- Do not blindly rewrite the single HTML file. Search and patch the relevant module section only.
- Preserve user-created templates and watermarks. The user explicitly dislikes updates overwriting templates.
- Avoid committing `dist/`, `node_modules/`, `_backups/`, or debug comparison folders.
- Before release, run:
  ```powershell
  node --check main.js
  node -e "const fs=require('fs');const s=fs.readFileSync('图片工作台.html','utf8');const m=s.match(/<script>([\s\S]*)<\/script>/);new Function(m?m[1]:'');console.log('html script ok')"
  ```
- Check `git status --short` before committing. At handoff time, local worktree may show uncommitted deletion of `templates/111/111/*` and an untracked `_backups/` directory. Those were intentionally not included in release commit `b75f476`.

## Current Known Follow-Ups
- Print AI UI is newly added and may need more UX refinement after real workflow use.
- Print AI image generation currently assumes OpenAI-compatible image endpoints. Packy `gpt-image-2` docs may require `/v1/images/generations`; check endpoint behavior before expanding.
- Product publish and print AI share some model configuration concepts but should stay separate unless intentionally unified.
- The app is a large single-file renderer. Any major refactor should first split modules carefully, with behavior snapshots/tests where possible.

