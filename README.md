# Markdown to PPTX Web

A privacy-friendly static Markdown / plain-text to editable PowerPoint converter with live slide preview, automatic pagination and GitHub Pages deployment.

## Features

- Paste Markdown or long Chinese/English text; preview slides automatically as you type.
- Large, editable PowerPoint text, with automatic page breaking and section headings.
- Download `.pptx` directly in the browser. Files are generated locally, without uploading source text.
- Import `.md` / `.txt`, select slide density, and use the sample document.
- Customize the default body font size (14–36 pt; 22 pt standard, 19 pt compact). The selected value controls both automatic pagination and editable PPTX text and is stored in this browser's localStorage (not uploaded).
- Adjust all four page margins independently (top / bottom / left / right, 0.5–5.0 cm). Preview and PPTX export use identical page geometry.
- Optionally display the document title at bottom left and slide numbers at bottom right. Both default to off; generated slides never include an MD2PPTX watermark.
- Preserve Markdown `**bold**` and `__bold__` emphasis within editable PowerPoint text runs, even after auto-pagination.
- Generate **one native editable text box per slide for non-table text**, including slide title, body, Markdown formatting, and optional footers. Markdown tables are separate **editable native PowerPoint table objects**, not text boxes or images. Decorative lines remain vector shapes.
- Parse GitHub-flavored Markdown pipe tables, including header rows, alignment markers, empty cells, inline bold, and escaped `\|`. Place short tables alongside preceding and following text on the same slide when space permits. Automatically paginate tall tables and repeat headers on continuation slides.
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

`src/pagination.mjs` contains a shared deterministic slide layout model used by both preview and PPTX export. The browser's font measurements drive line breaking, so rendered layout may differ slightly on devices with different fonts. Markdown headings, paragraphs, bold emphasis, lists, quotes, fenced code, and pipe tables are supported. Tables become native PowerPoint tables and can continue across slides, with repeated header rows. Merged cells, nested tables, and complex HTML table layouts are not supported; image references still become text. Unlike AI summarizers, this tool does not shorten your document.

The build bundles a pinned PptxGenJS distribution into the Pages artifact. The published site has no runtime CDN dependency. User documents remain in browser memory only and are not stored or sent to GitHub.

## Microsoft PowerPoint compatibility

Before any browser download, the generated PPTX is normalized to repair PptxGenJS 4.0.0's duplicate `p:cNvPr` object IDs, duplicate `a:pPr` elements within rich-text paragraphs, and incorrect table-frame extents. The workflow tests these OOXML invariants and opens representative documents in LibreOffice. Microsoft PowerPoint itself is not available on the Linux CI runner; its exact repair dialog cannot be asserted in CI.

## LaTeX equations

Inline `$...$` and `\\(...\\)`, as well as display `$$...$$` and `\\[...\\]` are rendered locally with a bundled MathJax SVG engine. PowerPoint embeds the equations as high-resolution PNG formula images for compatibility; surrounding prose remains native editable text. Equations themselves are **not native editable Office Math objects**, and future support for OMML requires separate validation. Unsupported LaTeX commands may trigger an error rendering placeholder.
