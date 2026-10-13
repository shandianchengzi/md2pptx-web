# Markdown to PPTX Web

A privacy-friendly static Markdown / plain-text to editable PowerPoint converter with live slide preview, automatic pagination and GitHub Pages deployment.

## Features

- Paste Markdown or long Chinese/English text; preview slides automatically as you type.
- Large, editable PowerPoint text, with automatic page breaking and section headings.
- Download `.pptx` directly in the browser. Files are generated locally, without uploading source text.
- Import `.md` / `.txt`, select slide density, and use the sample document.
- Customize the default body font size (14–36 pt; 22 pt standard, 19 pt compact). The selected value controls both automatic pagination and editable PPTX text and is stored in this browser's localStorage (not uploaded).
- GitHub Pages deployment through `.github/workflows/pages.yml`.

## Deploy

Set **Settings → Pages → Build and deployment → Source** to **GitHub Actions**. Push to `main` or run the `Deploy site` workflow manually. Expected project URL: https://shandianchengzi.github.io/md2pptx-web/ (requires Pages to be enabled).

## Local development

```bash
npm install
npm test
npm run build
npx http-server dist -p 8080
```

Open `http://localhost:8080`. Node 20+ recommended.

## Technical notes

`src/pagination.mjs` contains a shared deterministic slide layout model used by both preview and PPTX export. The browser's font measurements drive line breaking, so rendered layout may differ slightly on devices with different fonts. Basic Markdown (headings, paragraphs, bold/italic markers, lists, quotes, fenced code) is supported; complex tables/images are represented as text, not faithfully re-created as PPT graphics. Unlike AI summarizers, this tool does not shorten your document.

The build bundles a pinned PptxGenJS distribution into the Pages artifact. The published site has no runtime CDN dependency. User documents remain in browser memory only and are not stored or sent to GitHub.
