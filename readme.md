# PowerPoint → PDF: Export Illustrations for LaTeX

Export any PowerPoint illustration as a tightly cropped, vector-quality PDF on Windows. Text remains text, graphics remain vectors, no external tools required — no slide borders, no white margins, no rasterization.

![Preview](https://github.com/anoop-scholar/ppt-pdf-export/blob/main/ppt-pdf-export.png)
---

## Why?

PowerPoint is surprisingly effective for creating publication-quality scientific figures: flowcharts, model architectures, experimental pipelines, system diagrams. But getting those illustrations cleanly into LaTeX on Windows is awkward.

On a Mac, you right-click any shape or group and hit *Save as Picture → PDF*. The PDF is cropped exactly to the illustration's bounding box, vectors are preserved, text stays text.

**On Windows, that option doesn't exist.**

The common workarounds are:

- Export as PNG — rasterized, poor quality at print resolution
- Save the entire slide as PDF and crop later with `pdfcrop`
- Copy into Inkscape or Illustrator and re-export

Every time you update a figure, you repeat this process. For a paper with 8 figures going through 5 revision cycles, that is 40 manual crop operations. It adds up.

---

## The Solution

This VBA macro replicates the Mac right-click behaviour on Windows.

Select any shape, group, or illustration — regardless of whether it fits the slide, overflows it, or sits at an odd position — run the macro, and get a PDF cropped exactly to the illustration's bounding box, with all vectors and text preserved.

### How it works

1. A new blank presentation is created with its page sized **exactly** to the selected shape's bounding box
2. The shape is copied from your original file and pasted into the new presentation at position `(0, 0)`
3. The new presentation is exported as PDF — just your illustration on a perfectly-sized page
4. The temporary presentation is closed without saving

The result is a PDF whose page dimensions match your illustration exactly. Drop it straight into LaTeX with `\includegraphics`.

---

## Installation & Usage

1. Open your PowerPoint presentation and **save it** first
2. Press `Alt + F11` to open the VBA editor
3. Click **Insert → Module** — a white code editor appears on the right
4. Open [`ExportIllustrationAsPDF.txt`](ExportIllustrationAsPDF.txt), copy all contents, and paste into the module
5. Close the VBA editor (`Alt + F4`)
6. Back in PowerPoint, **select your illustration** — a single shape, a group, or anything you've built
7. Press `Alt + F8`, select `ExportIllustrationAsPDF`, and click **Run**
8. The PDF is saved to the same folder as your `.pptx`, named after the shape

> **Tip:** Add the macro to your Quick Access Toolbar (File → Options → Quick Access Toolbar → Macros) so it's a single click — just like the Mac right-click.

> **Note:** When you paste the macro, PowerPoint will show a yellow bar saying *"VBA projects must be saved in a macro-enabled presentation"* — click **Save As** and save as `.pptm` to keep the macro. When closing, if asked whether to save as `.pptm` or `.pptx`, choose `.pptm` to retain the macro for future use, or `.pptx` to discard it — you can always re-paste it from this page.

---

## Requirements

- Windows
- Microsoft PowerPoint (any modern version)
- Microsoft Office VBA enabled (default in most installations)

---
### Files Included

- `ExportIllustrationAsPDF.bas` — standard VBA module. You can import it directly into the VBA editor via File → Import File….
- `ExportIllustrationAsPDF.txt` — plain-text copy for users who prefer to open the file and copy-paste the code manually.
---

## Why not just use `pdfcrop`?

`pdfcrop` is a valid workaround but adds a step to every figure iteration. It also requires a LaTeX distribution installed and invoked from the command line. This macro does the same thing in one click, from inside PowerPoint, with no external tools.

---

## Contributing

If you run into edge cases — illustrations with embedded images, linked files, or unusual shape types — please open an issue or PR.
