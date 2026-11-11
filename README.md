# DOCX to Markdown Converter

A small, dependency-light Python program that converts Word documents (`.docx`) into Markdown (`.md`). Run it, enter the input and output paths when prompted, and you get a clean Markdown file.

## Features

- Headings (`Heading 1`–`Heading 6`) and the Title style
- **Bold**, *italic*, ***bold-italic***, ~~strikethrough~~ and `inline code` (monospace fonts)
- Hyperlinks
- Bullet and numbered lists, including nested lists
- Block quotes
- Tables (converted to Markdown tables)
- Images, extracted into a folder next to the output file and linked automatically
- Interactive prompts with input validation (extension checks, file existence, overwrite confirmation)
- Footnote support
- Code block detection
- Command-line arguments as an alternative to the interactive prompts

## Requirements

- Python 3.8 or newer (Python 3.12+ recommended)
- [python-docx](https://pypi.org/project/python-docx/)

## Installation

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install python-docx
```

## Usage

```bash
python docx_to_md.py
```

The program asks for two paths. Both must include the file extension:

```
=== DOCX to Markdown Converter ===

Full path of the .docx file (including extension): C:\Users\me\Documents\report.docx
Full output path (including extension, e.g. C:\out\result.md): C:\Users\me\Documents\report.md

Done! Markdown saved to: C:\Users\me\Documents\report.md
1 image(s) saved to: C:\Users\me\Documents\report_images
```

Paths pasted with surrounding quotes (for example from Windows "Copy as path") are handled automatically.

### Images

If the document contains images, they are saved in a folder named after the output file with `_images` appended (for example `report.md` → `report_images/`) and referenced in the Markdown with relative links. Keep the `.md` file and its image folder together when moving or uploading them.

## Example

**Word document**

> # Main Title
> Normal text with **bold**, *italic* and a [link](https://example.com).
> - Bullet A
> - Bullet B
>   - Nested
>
> | Name | Value |
> |------|-------|
> | x    | 2     |

**Resulting Markdown**

```markdown
# Main Title

Normal text with **bold**, *italic* and a [link](https://example.com).

- Bullet A
- Bullet B
    - Nested

| Name | Value |
| --- | --- |
| x | 2 |
```

## Limitations

Markdown is much simpler than Word, so some things are simplified or dropped:

- Colors, fonts, font sizes and text alignment
- Footnotes, endnotes, headers and footers
- Text boxes, shapes and charts
- Merged table cells (the content is kept, the merge is not)
- Only `.docx` files are supported (not legacy `.doc`)

## How it works

The script reads the document with `python-docx` and walks through the body in order, converting paragraphs and tables one by one. Adjacent text runs with identical formatting are merged to avoid output like `**a****b**`. List types (bullet or numbered) are read from the document's numbering definitions.

###### Created by wkg2673
