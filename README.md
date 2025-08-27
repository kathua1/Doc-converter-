# Automatic Document Converter - Static Site

This is a simple static website that provides client-side document conversions:
- Image → PDF (jsPDF)
- PDF → Images (pdf.js)
- PDF → Word (.docx) — inserts page images into the Word file
- PDF → Excel (.xlsx) — extracts page text into rows
- DOCX → PDF (mammoth + jsPDF)
- Image(s) → PowerPoint (.pptx) (PptxGenJS)

Notes:
- All conversions run in the user's browser — no files are uploaded to servers.
- Some conversions are approximate; complex PDFs with tables/layouts may not be perfectly preserved.
- You can host this on GitHub Pages. Upload the contents of this repo (or the generated ZIP) and enable Pages.

## How to use
1. Open `index.html` in a modern browser (Chrome/Edge/Firefox).
2. Choose files and click convert.
3. Download the generated file.

## Libraries (CDN)
The site uses popular CDN-hosted libraries. For offline use, replace CDNs with local copies.

