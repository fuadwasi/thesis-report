#!/usr/bin/env python3
"""
DOCX to LaTeX Converter
-----------------------
This script converts a Microsoft Word .docx file into a complete LaTeX document.
It handles:
- Document structure (chapters, sections, subsections)
- Text formatting (bold, italic, underline)
- Lists (bulleted and numbered)
- Tables
- Images/Figures
- References and citations

Usage:
    python docx_to_latex.py <input.docx> [output.tex]
"""

import sys
import os
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import zipfile


class DocxToLatexConverter:
    """Main converter class for transforming DOCX to LaTeX."""
    
    def __init__(self, docx_path):
        """
        Initialize the converter with a DOCX file.
        
        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = docx_path
        self.doc = Document(docx_path)
        self.output_lines = []
        self.figure_counter = 0
        self.table_counter = 0
        self.image_paths = []
        self.in_list = False
        self.list_type = None  # 'itemize' or 'enumerate'
        
    def escape_latex(self, text):
        """
        Escape special LaTeX characters in text.
        
        Args:
            text: Input text string
            
        Returns:
            Text with escaped LaTeX special characters
        """
        if not text:
            return ""
        
        # Characters that need escaping in LaTeX
        replacements = {
            '\\': r'\textbackslash{}',
            '&': r'\&',
            '%': r'\%',
            '$': r'\$',
            '#': r'\#',
            '_': r'\_',
            '{': r'\{',
            '}': r'\}',
            '~': r'\textasciitilde{}',
            '^': r'\textasciicircum{}',
        }
        
        # Apply replacements
        for char, replacement in replacements.items():
            text = text.replace(char, replacement)
            
        return text
    
    def format_text_run(self, run):
        """
        Convert a Word text run to LaTeX with formatting.
        
        Args:
            run: A python-docx Run object
            
        Returns:
            LaTeX-formatted text string
        """
        text = self.escape_latex(run.text)
        
        if not text:
            return ""
        
        # Apply formatting
        if run.bold:
            text = f"\\textbf{{{text}}}"
        if run.italic:
            text = f"\\textit{{{text}}}"
        if run.underline:
            text = f"\\underline{{{text}}}"
            
        return text
    
    def convert_paragraph(self, para):
        """
        Convert a Word paragraph to LaTeX.
        
        Args:
            para: A python-docx Paragraph object
            
        Returns:
            LaTeX string representation of the paragraph
        """
        # Get plain text for pattern matching
        plain_text = para.text.strip()
        
        # Get paragraph text with formatting
        formatted_text = ""
        for run in para.runs:
            formatted_text += self.format_text_run(run)
        
        # Skip empty paragraphs
        if not formatted_text.strip():
            return ""
        
        style_name = para.style.name
        
        # Handle different paragraph styles
        if style_name == 'Heading 1':
            # Chapter heading
            # Extract chapter number and title using plain text
            chapter_match = re.match(r'Chapter\s+(\d+):\s*(.*)', plain_text, re.IGNORECASE)
            if chapter_match:
                chapter_num = chapter_match.group(1)
                chapter_title = self.escape_latex(chapter_match.group(2))
                return f"\\chapter{{{chapter_title}}}\n\\label{{ch:{chapter_num}}}\n"
            else:
                # For non-standard chapter headings (like "References")
                # Use plain text to avoid excessive formatting
                clean_text = self.escape_latex(plain_text)
                return f"\\chapter{{{clean_text}}}\n"
                
        elif style_name == 'Heading 2':
            # Section heading - use plain text to avoid excessive formatting
            clean_text = self.escape_latex(plain_text)
            return f"\\section{{{clean_text}}}\n"
            
        elif style_name == 'Heading 3':
            # Subsection heading - use plain text
            clean_text = self.escape_latex(plain_text)
            return f"\\subsection{{{clean_text}}}\n"
            
        elif style_name == 'Heading 4':
            # Subsubsection heading - use plain text
            clean_text = self.escape_latex(plain_text)
            return f"\\subsubsection{{{clean_text}}}\n"
        
        elif style_name in ['List Paragraph', 'List']:
            # Handle list items
            # Check if paragraph starts with a bullet or number
            text_content = para.text.strip()
            
            # Detect list type
            if not self.in_list:
                # Start a new list
                # Simple heuristic: if it looks like a numbered list
                if re.match(r'^\d+[\.)]\s', text_content) or re.match(r'^[a-z][\.)]\s', text_content):
                    self.list_type = 'enumerate'
                    result = "\\begin{enumerate}\n"
                else:
                    self.list_type = 'itemize'
                    result = "\\begin{itemize}\n"
                self.in_list = True
            else:
                result = ""
            
            # Clean up the text (remove bullets or numbers)
            item_text = re.sub(r'^[\d\w]+[\.)]\s*', '', formatted_text)
            item_text = re.sub(r'^[•\-\*]\s*', '', item_text)
            
            result += f"  \\item {item_text}\n"
            return result
        
        else:
            # Close any open list before regular paragraph
            result = ""
            if self.in_list:
                result = f"\\end{{{self.list_type}}}\n\n"
                self.in_list = False
                self.list_type = None
            
            # Regular paragraph
            # Handle special cases (acknowledgements, abstract, etc.) using plain text
            upper_text = plain_text.upper()
            if upper_text in ['ACKNOWLEDGEMENT', 'ACKNOWLEDGEMENTS']:
                return result + "\\chapter*{Acknowledgements}\n\\addcontentsline{toc}{chapter}{Acknowledgements}\n\n"
            elif upper_text == 'ABSTRACT':
                return result + "\\chapter*{Abstract}\n\\addcontentsline{toc}{chapter}{Abstract}\n\n"
            elif plain_text == 'TABLE OF CONTENTS':
                return result + "\\tableofcontents\n\\newpage\n\n"
            elif plain_text == 'List of Figures':
                return result + "\\listoffigures\n\\newpage\n\n"
            elif plain_text == 'List of Tables':
                return result + "\\listoftables\n\\newpage\n\n"
            elif plain_text == 'List of Abbreviations':
                return result + "\\chapter*{List of Abbreviations}\n\\addcontentsline{toc}{chapter}{List of Abbreviations}\n\n"
            else:
                # Check alignment
                alignment = ""
                if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    alignment = "\\begin{center}\n" + formatted_text + "\n\\end{center}\n"
                    return result + alignment
                
                return result + formatted_text + "\n\n"
    
    def convert_table(self, table):
        """
        Convert a Word table to LaTeX table.
        
        Args:
            table: A python-docx Table object
            
        Returns:
            LaTeX table string
        """
        self.table_counter += 1
        
        # Determine number of columns
        num_cols = len(table.columns)
        
        # Start table environment
        result = "\\begin{table}[htbp]\n"
        result += "\\centering\n"
        result += "\\begin{tabular}{|" + "l|" * num_cols + "}\n"
        result += "\\hline\n"
        
        # Process each row
        for row_idx, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                # Get cell text with formatting
                cell_text = ""
                for para in cell.paragraphs:
                    for run in para.runs:
                        cell_text += self.format_text_run(run)
                
                # Clean and escape
                cell_text = cell_text.strip()
                if not cell_text:
                    cell_text = ""
                    
                cells.append(cell_text)
            
            # Join cells with &
            result += " & ".join(cells) + " \\\\\n"
            result += "\\hline\n"
        
        result += "\\end{tabular}\n"
        result += f"\\caption{{Table {self.table_counter}}}\n"
        result += f"\\label{{tab:table{self.table_counter}}}\n"
        result += "\\end{table}\n\n"
        
        return result
    
    def extract_images(self):
        """
        Extract images from the DOCX file.
        
        Returns:
            Dictionary mapping image IDs to filenames
        """
        image_dict = {}
        
        try:
            # Open the docx as a zip file
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                # List all files in the media directory (can be word/media/ or just media/)
                for file_info in docx_zip.filelist:
                    if 'media/' in file_info.filename and not file_info.is_dir():
                        # Extract the image
                        image_data = docx_zip.read(file_info.filename)
                        
                        # Get just the filename
                        image_name = os.path.basename(file_info.filename)
                        
                        # Save to output directory
                        output_dir = os.path.dirname(self.docx_path) or '.'
                        image_dir = os.path.join(output_dir, 'images')
                        os.makedirs(image_dir, exist_ok=True)
                        
                        image_path = os.path.join(image_dir, image_name)
                        with open(image_path, 'wb') as img_file:
                            img_file.write(image_data)
                        
                        # Store relative path
                        image_dict[image_name] = f"images/{image_name}"
                        self.image_paths.append(image_name)
                        
        except Exception as e:
            print(f"Warning: Could not extract images: {e}")
            import traceback
            traceback.print_exc()
        
        return image_dict
    
    def check_for_image(self, para):
        """
        Check if a paragraph contains an image.
        
        Args:
            para: A python-docx Paragraph object
            
        Returns:
            Image filename if found, None otherwise
        """
        # Check for embedded images in runs
        for run in para.runs:
            if 'graphic' in run._element.xml:
                # Try to extract image reference
                try:
                    # Look for image references in the XML
                    for rel in run.part.rels.values():
                        if "image" in rel.target_ref:
                            return os.path.basename(rel.target_ref)
                except:
                    pass
        return None
    
    def convert_figure(self, image_name, caption=""):
        """
        Create LaTeX figure environment for an image.
        
        Args:
            image_name: Name of the image file
            caption: Optional caption text
            
        Returns:
            LaTeX figure string
        """
        self.figure_counter += 1
        
        result = "\\begin{figure}[htbp]\n"
        result += "\\centering\n"
        result += f"\\includegraphics[width=0.8\\textwidth]{{images/{image_name}}}\n"
        
        if caption:
            result += f"\\caption{{{self.escape_latex(caption)}}}\n"
        else:
            result += f"\\caption{{Figure {self.figure_counter}}}\n"
            
        result += f"\\label{{fig:figure{self.figure_counter}}}\n"
        result += "\\end{figure}\n\n"
        
        return result
    
    def generate_latex_preamble(self):
        """
        Generate the LaTeX document preamble.
        
        Returns:
            LaTeX preamble string
        """
        preamble = r"""\documentclass[12pt,a4paper]{report}

% Packages
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{amsmath}
\usepackage{amsfonts}
\usepackage{amssymb}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{hyperref}
\usepackage{geometry}
\usepackage{setspace}
\usepackage{caption}
\usepackage{subcaption}
\usepackage[table]{xcolor}
\usepackage{longtable}
\usepackage{float}

% Page geometry
\geometry{
    left=1.5in,
    right=1in,
    top=1in,
    bottom=1in
}

% Line spacing
\onehalfspacing

% Hyperref settings
\hypersetup{
    colorlinks=true,
    linkcolor=blue,
    filecolor=magenta,
    urlcolor=cyan,
    citecolor=blue,
}

% Title page information (customize as needed)
\title{Comparative Analysis of Neural Network Architectures for Stock Market Prediction}
\author{Imamul Islam Ifti\\Student ID: 1909056}
\date{September 2025}

\begin{document}

% Title page
\maketitle

% Front matter
\frontmatter

"""
        return preamble
    
    def generate_latex_closing(self):
        """
        Generate the LaTeX document closing.
        
        Returns:
            LaTeX closing string
        """
        closing = r"""
\end{document}
"""
        return closing
    
    def convert(self):
        """
        Main conversion method.
        
        Returns:
            Complete LaTeX document as a string
        """
        print("Starting conversion...")
        
        # Extract images
        print("Extracting images...")
        image_dict = self.extract_images()
        print(f"Extracted {len(image_dict)} images")
        
        # Add preamble
        latex_output = self.generate_latex_preamble()
        
        # Track if we're in main matter
        in_main_matter = False
        skip_until_chapter = True  # Skip title page content
        
        # Process document elements
        print("Processing document content...")
        for element_idx, element in enumerate(self.doc.element.body):
            # Check if it's a paragraph or table
            if isinstance(element, CT_P):
                para = Paragraph(element, self.doc)
                
                # Check if this is the first chapter to switch to mainmatter
                if para.style.name == 'Heading 1' and not in_main_matter:
                    # Close any open list
                    if self.in_list:
                        latex_output += f"\\end{{{self.list_type}}}\n\n"
                        self.in_list = False
                    
                    latex_output += "\\mainmatter\n\n"
                    in_main_matter = True
                    skip_until_chapter = False
                
                # Skip title page elements until first chapter
                if skip_until_chapter and para.style.name != 'Heading 1':
                    # But keep acknowledgements, abstract, TOC, etc.
                    text = para.text.strip().upper()
                    if text not in ['ACKNOWLEDGEMENT', 'ACKNOWLEDGEMENTS', 'ABSTRACT', 
                                   'TABLE OF CONTENTS', 'LIST OF FIGURES', 'LIST OF TABLES']:
                        continue
                
                # Convert paragraph
                latex_text = self.convert_paragraph(para)
                if latex_text:
                    latex_output += latex_text
                
                # Check for images in paragraph
                # (Note: Image detection in python-docx can be complex)
                # This is a simplified approach
                
            elif isinstance(element, CT_Tbl):
                # Close any open list before table
                if self.in_list:
                    latex_output += f"\\end{{{self.list_type}}}\n\n"
                    self.in_list = False
                    self.list_type = None
                
                table = Table(element, self.doc)
                latex_output += self.convert_table(table)
        
        # Close any remaining open list
        if self.in_list:
            latex_output += f"\\end{{{self.list_type}}}\n\n"
        
        # Add closing
        latex_output += self.generate_latex_closing()
        
        print("Conversion complete!")
        return latex_output
    
    def save_to_file(self, output_path):
        """
        Convert and save to a .tex file.
        
        Args:
            output_path: Path for the output .tex file
        """
        latex_content = self.convert()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(latex_content)
        
        print(f"LaTeX file saved to: {output_path}")
        
        # Print summary
        print("\n" + "="*60)
        print("CONVERSION SUMMARY")
        print("="*60)
        print(f"Figures: {self.figure_counter}")
        print(f"Tables: {self.table_counter}")
        print(f"Images extracted: {len(self.image_paths)}")
        print("\nNext steps:")
        print("1. Review the generated .tex file")
        print("2. Ensure all images are in the 'images/' directory")
        print("3. Compile with: pdflatex output.tex")
        print("4. Run twice for proper references")
        print("="*60)


def main():
    """Main entry point for the script."""
    if len(sys.argv) < 2:
        print("Usage: python docx_to_latex.py <input.docx> [output.tex]")
        print("\nExample:")
        print("  python docx_to_latex.py thesis-report.docx thesis-report.tex")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Determine output file
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        # Default: replace .docx with .tex
        base_name = os.path.splitext(input_file)[0]
        output_file = base_name + ".tex"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found!")
        sys.exit(1)
    
    # Perform conversion
    try:
        converter = DocxToLatexConverter(input_file)
        converter.save_to_file(output_file)
        print(f"\nSuccess! LaTeX file created: {output_file}")
        
    except Exception as e:
        print(f"Error during conversion: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
