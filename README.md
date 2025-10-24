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
