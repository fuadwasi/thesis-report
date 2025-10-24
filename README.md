# DOCX to LaTeX Converter

This repository contains a Python script to convert Microsoft Word (.docx) files to LaTeX (.tex) format.

## Features

- **Document Structure Conversion**: Automatically converts Word headings to LaTeX chapters, sections, and subsections
- **Text Formatting Preservation**: Maintains bold, italic, and underline formatting
- **List Support**: Converts both bulleted (itemize) and numbered (enumerate) lists
- **Table Conversion**: Transforms Word tables into LaTeX tabular environments
- **Image Extraction**: Extracts embedded images from the .docx file to an `images/` directory
- **Figure Support**: Creates LaTeX figure environments for images
- **Special Sections**: Handles front matter like Abstract, Acknowledgements, Table of Contents, List of Figures, and List of Tables
- **Clean LaTeX Output**: Produces well-formatted, compilable LaTeX code

## Requirements

- Python 3.6 or higher
- python-docx library

## Installation

1. Install the required Python library:

```bash
pip install python-docx
```

## Usage

### Basic Usage

Convert a .docx file to LaTeX:

```bash
python docx_to_latex.py thesis-report.docx
```

This will create `thesis-report.tex` in the same directory.

### Specify Output File

You can specify a custom output filename:

```bash
python docx_to_latex.py input.docx output.tex
```

### Example

```bash
python docx_to_latex.py thesis-report.docx thesis-report.tex
```

This will:
1. Extract all images to the `images/` directory
2. Convert the document structure to LaTeX
3. Create `thesis-report.tex` with proper LaTeX formatting

## Output

The script generates:
- A `.tex` file with the complete LaTeX document
- An `images/` directory containing all extracted images

## Compiling the LaTeX Document

After conversion, compile the LaTeX document with:

```bash
pdflatex thesis-report.tex
pdflatex thesis-report.tex  # Run twice for proper references
```

Or use your preferred LaTeX editor (TeXworks, TeXmaker, Overleaf, etc.)

## Document Structure

The converter handles the following Word styles:

- **Heading 1** → `\chapter{}`
- **Heading 2** → `\section{}`
- **Heading 3** → `\subsection{}`
- **Heading 4** → `\subsubsection{}`
- **List Paragraph** → `\item` in `itemize` or `enumerate` environment
- **Normal** → Regular paragraphs

### Special Sections

The script recognizes and properly formats:
- ACKNOWLEDGEMENT → `\chapter*{Acknowledgements}`
- ABSTRACT → `\chapter*{Abstract}`
- TABLE OF CONTENTS → `\tableofcontents`
- List of Figures → `\listoffigures`
- List of Tables → `\listoftables`

## LaTeX Document Class

The generated LaTeX document uses:
- Document class: `report`
- Paper size: A4
- Font size: 12pt
- Line spacing: 1.5 (onehalfspacing)

## Included LaTeX Packages

The generated document includes:
- `graphicx` - For images
- `booktabs` - For better tables
- `hyperref` - For clickable links and references
- `amsmath`, `amsfonts`, `amssymb` - For mathematical symbols
- `geometry` - For page layout
- `caption` and `subcaption` - For figure captions
- `longtable` - For tables spanning multiple pages
- `float` - For better float positioning

## Limitations

- Complex Word formatting may not convert perfectly
- Embedded equations are not automatically converted (manual adjustment may be needed)
- Comments and track changes are not preserved
- Some advanced Word features may require manual LaTeX adjustment

## Customization

You can customize the generated LaTeX document by editing:
- The preamble in `generate_latex_preamble()` method
- Page geometry settings
- Font and spacing options
- Package inclusions

## Example Workflow

1. Place your `.docx` file in the repository directory
2. Run the converter: `python docx_to_latex.py yourfile.docx`
3. Review the generated `.tex` file
4. Check that images are properly extracted to `images/` directory
5. Compile with LaTeX: `pdflatex yourfile.tex`
6. Make any necessary manual adjustments
7. Compile again for proper cross-references

## Troubleshooting

### Images not showing
- Ensure the `images/` directory exists and contains the extracted images
- Check that the image paths in the `.tex` file are correct
- Verify that your LaTeX distribution supports the image formats (PNG, JPEG)

### Compilation errors
- Run `pdflatex` twice to resolve cross-references
- Check for special characters that need escaping
- Review any complex tables or formatting that may need manual adjustment

### Missing content
- Check that the original .docx file is not corrupted
- Ensure all required styles (Heading 1, Heading 2, etc.) are used correctly in Word

## License

This script is provided as-is for educational and research purposes.

## Contributing

Feel free to submit issues or pull requests for improvements.

## Sample Output

When you run the converter, you'll see output like this:

```
Starting conversion...
Extracting images...
Extracted 26 images
Processing document content...
Conversion complete!
LaTeX file saved to: thesis-report.tex

============================================================
CONVERSION SUMMARY
============================================================
Figures: 0
Tables: 9
Images extracted: 26

Next steps:
1. Review the generated .tex file
2. Ensure all images are in the 'images/' directory
3. Compile with: pdflatex output.tex
4. Run twice for proper references
============================================================

Success! LaTeX file created: thesis-report.tex
```

## Features Demonstrated

### Heading Conversion
- Word "Heading 1" style with "Chapter 1: Introduction" becomes `\chapter{Introduction}\label{ch:1}`
- Word "Heading 2" becomes `\section{}`
- Word "Heading 3" becomes `\subsection{}`

### Text Formatting
- **Bold text** in Word becomes `\textbf{Bold text}`
- *Italic text* becomes `\textit{Italic text}`
- Underlined text becomes `\underline{Underlined text}`

### Lists
Word lists are converted to LaTeX itemize/enumerate:
```latex
\begin{itemize}
  \item First item
  \item Second item
\end{itemize}
```

### Tables
Word tables become LaTeX tabular environments with proper formatting and captions.

## Project Structure

After running the converter:
```
thesis-report/
├── thesis-report.docx       # Original Word document
├── thesis-report.tex        # Generated LaTeX document
├── docx_to_latex.py         # Conversion script
├── images/                  # Extracted images
│   ├── image.png
│   ├── image2.png
│   └── ...
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Advanced Usage

### Command Line Options
```bash
# Display help
python docx_to_latex.py --help

# Basic conversion (output will be thesis-report.tex)
python docx_to_latex.py thesis-report.docx

# Specify output file
python docx_to_latex.py input.docx custom-output.tex
```

## Technical Details

### Supported Word Styles
- Heading 1, 2, 3, 4
- Normal paragraphs
- List Paragraph
- Centered text
- Bold, Italic, Underline formatting

### LaTeX Packages Used
The generated document includes standard packages for academic writing:
- Graphics and figures: `graphicx`, `float`
- Tables: `booktabs`, `longtable`
- Math: `amsmath`, `amsfonts`, `amssymb`
- Cross-references: `hyperref`
- Page layout: `geometry`, `setspace`

### Character Escaping
Special LaTeX characters are automatically escaped:
- `&` → `\&`
- `%` → `\%`
- `$` → `\$`
- `#` → `\#`
- `_` → `\_`
- `{`, `}` → `\{`, `\}`
- `\` → `\textbackslash{}`

## Testing

A simple validation can be run to ensure the conversion works:

```python
import os

# Check if files exist
assert os.path.exists('thesis-report.tex')
assert os.path.exists('images/')

# Validate LaTeX content
with open('thesis-report.tex', 'r') as f:
    content = f.read()
    assert r'\documentclass' in content
    assert r'\chapter{' in content
    assert r'\begin{document}' in content
    assert r'\end{document}' in content
```

## FAQ

**Q: Can I use this with other document formats?**
A: This script is specifically designed for .docx files. For other formats, consider using Pandoc.

**Q: What if my document has complex equations?**
A: Mathematical equations in Word may not convert perfectly. You may need to manually adjust them in LaTeX.

**Q: How do I handle citations and bibliography?**
A: This script converts the References section as regular text. For proper BibTeX integration, you'll need to manually create a .bib file and update the LaTeX accordingly.

**Q: Can I customize the LaTeX output?**
A: Yes! Edit the `generate_latex_preamble()` method in the script to customize packages, page layout, and other settings.

## Version History

- **v1.0** (2025-10-24): Initial release with full conversion support for headings, formatting, lists, tables, and images

## Author

Created for converting academic thesis documents from Word to LaTeX format.
