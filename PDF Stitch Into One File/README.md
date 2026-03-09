# PDF Stitch Into One File

A small offline tool that merges PDFs and images into one PDF in the order you add them.

## Usage
1. Open `index.html` in a modern browser.
2. Drag and drop PDF, JPG, or PNG files or use "Choose Files".
3. Choose the image page size if you want to override the default.
4. Use "PDF page fit" to trim margins when crop or trim boxes are present.
5. Add watermark text if you want a licensing mark on selected pages.
6. Adjust watermark opacity, tone, size, orientation, placement, and preview orientation in the live preview.
7. Use the watermark pages list to pick where the watermark appears.
8. Add footer text if you want a footer on selected pages.
9. Adjust footer opacity and tone, then pick the footer pages.
10. Click "Stitch to PDF" and download the merged PDF.

## Notes
- Everything runs locally in your browser; no uploads.
- Drag items in the queue to reorder the merge sequence.
- Images are scaled to fit the selected image page size (defaults to first PDF page, or Letter when no PDF is present).
- Watermark text can be diagonal, straight, or inverse and scales with the size control.
- PDF page fit can trim extra margins if the source PDF defines crop or trim boxes.
- Footer text is centered at the bottom of selected pages (no box) with adjustable opacity and tone.
- Watermark preview swatches help tune the watermark for light and dark content.
- Drag the watermark box in the preview to reposition and resize it.
- Footer and watermark page checkboxes update automatically as files are added or removed.
- The included `vendor/pdf-lib.min.js` comes from pdf-lib (MIT license).
