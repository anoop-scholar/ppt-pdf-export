# PowerPoint → PDF: Export Illustrations for LaTeX

Export any illustration from PowerPoint as a **tight-cropped, vector-quality PDF** — on Windows. No slide borders, no white margins, no rasterization.

---

## The Problem

If you write papers in LaTeX, you know the rule: **always use PDF figures**. Vector graphics scale perfectly, render crisply in print, and keep file sizes small. The moment you drop a PNG or JPEG into your `.tex` file, you've committed to a blurry figure at any resolution that wasn't exactly what you exported at.

PowerPoint is, despite everything, a genuinely good tool for making scientific illustrations. Flowcharts, system diagrams, model architectures, experimental pipelines — the combination of aligned shapes, editable text, and freeform drawing is hard to beat for quick, publication-quality figures.

**But getting those illustrations into LaTeX on Windows is painful.**

On a Mac, you right-click any shape or group and hit *Save as Picture → PDF*. Done. The PDF is cropped exactly to the illustration's bounding box, vectors are preserved, text stays text.

On Windows, that option doesn't exist.

Your options, as a Windows researcher, are:
- **Export as PNG** — rasterized, not suitable for LaTeX figures at print resolution
- **Save as PDF** — exports the entire slide, leaving you with enormous white margins you have to crop manually with `pdfcrop` or Adobe Acrobat
- **Use Inkscape or Illustrator** — copy-paste from PowerPoint, fix broken styles, re-export. Tedious for every figure iteration.
- **Switch to a Mac** — not always an option

Every time you update a figure, you repeat this process. For a paper with 8 figures that each go through 5 revision cycles, that is 40 manual crop operations. It adds up.

---

## The Solution

This VBA macro replicates the Mac right-click behaviour on Windows.

Select any shape, group, or illustration in PowerPoint — regardless of whether it fits the slide, overflows it, or sits at an odd position — run the macro, and get a PDF cropped exactly to the illustration's bounding box, with all vectors and text preserved.

### How it works

The technique avoids all the canvas-resizing hacks that cause fiddly side effects. Instead:

1. A new blank presentation is created invisibly, with its page sized **exactly** to the selected shape's bounding box
2. The shape is copied from your original file and pasted into the new presentation at position `(0, 0)`
3. The new presentation is exported as PDF — which is now just your illustration on a perfectly-sized page
4. The temporary presentation is closed without saving

The result is a PDF whose page dimensions match your illustration exactly. Drop it straight into LaTeX with `\includegraphics`.

---

## Installation

1. Open your PowerPoint presentation
2. Press `Alt + F11` to open the VBA editor
3. Click **Insert → Module**
4. Paste the contents of [`ExportIllustrationAsPDF.bas`](ExportIllustrationAsPDF.bas) into the module
5. Close the VBA editor

**To add it to your Quick Access Toolbar** (recommended — makes it one click):
- File → Options → Quick Access Toolbar
- Choose *Macros* from the "Choose commands from" dropdown
- Add `ExportIllustrationAsPDF` to the toolbar

---

## Usage

1. Select your illustration (a shape, group, or any combination)
2. Run the macro via `Alt + F8 → ExportIllustrationAsPDF → Run`, or click the toolbar button if you added it
3. The PDF is saved to the same folder as your `.pptx`, named after the shape

If your presentation hasn't been saved yet, the PDF goes to your Desktop.

---

## LaTeX usage

```latex
\begin{figure}[t]
    \centering
    \includegraphics[width=\linewidth]{figures/my_diagram.pdf}
    \caption{Your caption here.}
    \label{fig:my_diagram}
\end{figure}
```

Because the output is a true vector PDF, it scales to any `width` without quality loss.

---

## Requirements

- Windows
- Microsoft PowerPoint (any modern version)
- Microsoft Office VBA enabled (default in most installations)

---

## Why not just use `pdfcrop`?

`pdfcrop` is a valid workaround but adds a step to every iteration of every figure. It also requires a LaTeX distribution to be installed and invoked from the command line. This macro does the same thing in one click, from inside PowerPoint, with no external tools.

---

## Contributing

If you find edge cases — illustrations with embedded videos, linked images, or unusual shape types — please open an issue or PR.
