# PowerPoint → PDF: Export Illustrations for LaTeX

Export PowerPoint illustrations as tightly cropped, vector-quality PDFs on Windows. Text remains text, graphics remain vectors, and no external tools are required. No slide borders, no white margins, no rasterization.

---

## Why?

PowerPoint is surprisingly effective for creating quick, publication-quality scientific figures, flowcharts, model architectures, and experimental pipelines. However, exporting those illustrations cleanly for LaTeX on Windows is awkward.

### Note:
On a Mac, you right-click any shape or group and hit *Save as Picture → PDF*. Done. The PDF is cropped exactly to the illustration's bounding box, vectors are preserved, text stays text.

On Windows, that option doesn't exist.

### Common workarounds include:

* Exporting PNGs (rasterized)
* Saving the entire slide as PDF and cropping later
* Using pdfcrop
* Copying into Inkscape or Illustrator and re-exporting

These approaches add unnecessary friction to what should be a one-click operation.
Every time you update a figure, you repeat this process. For a paper with 8 figures that each go through 5 revision cycles, that is 40 manual crop operations. It adds up.

If you create diagrams for LaTeX documents, vector formats such as PDF are often preferable. They scale cleanly, preserve text and line quality, and typically produce smaller files than high-resolution raster images.
---

## The Solution

This VBA macro replicates the Mac right-click behaviour on Windows.

Select any shape, group, or illustration in PowerPoint — regardless of whether it fits the slide, overflows it, or sits at an odd position — run the macro, and get a PDF cropped exactly to the illustration's bounding box, with all vectors and text preserved.
---
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
