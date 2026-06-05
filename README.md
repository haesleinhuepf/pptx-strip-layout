# pptx-strip-layout

`pptx-strip-layout` is a pip-installable CLI utility that processes a `.pptx` file to:

- Remove images and non-placeholder objects from slide templates (masters and layouts)
- Set slide/template backgrounds to white
- Set text color to black in templates and slides
- Optionally convert slide images to grayscale and replace originals

## Installation

```bash
pip install .
```

## Usage

```bash
pptx-strip-layout input.pptx
```

Optional grayscale image conversion on slides:

```bash
pptx-strip-layout input.pptx --grayscale-images
```

Output file is saved as `input-wolayout.pptx` in the same folder.
