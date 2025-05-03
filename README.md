# Clause Inserter
A Python tool for programmatically inserting clauses into legal documents with precise formatting control, built using the [python-docx](https://python-docx.readthedocs.io/en/latest/) library.

## Overview
This tool allows users to insert clauses, sentences, or paragraphs into Word documents while maintaining consistent formatting, numbering, and styling. Built using the python-docx library. It features an interactive CLI interface.

## Features
- CLI Interactive interface
- Insert clauses before or after paragraphs
- Insert text between sentences in a paragraph
- Maintain document numbering and list styles
- Preserve formatting of surrounding text
- Interactive paragraph selection

## Getting Started

### Prerequisites
- Python 3.6 or higher
- Word documents (.docx format)

### Installation

```bash
# Clone the repository
git clone git@github.com:cdznd/clause-inserter.git
cd clause-inserter

# Set up virtual environment
python -m venv venv

# Activate virtual environment
## On macOS/Linux
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

Run the tool from the command line by providing a Word document path:

```bash
python clause_inserter.py path/to/document.docx
```

The tool will guide you through an interactive process to:
1. Select a target paragraph
2. Choose insertion position (before, after, or between sentences)
3. Configure formatting options
4. Insert your clause text

## Technologies
- **Python**: Core programming language
- **python-docx**: Library for working with Word documents
- **haggis**: Supporting library
- **lxml**: XML processing library
- **numpy**: Numerical computing library
- **CLI**: Command-line interface for user interaction

## Project Structure
```
src/
├── clause_inserter.py  # Main application file
├── cli.py              # Command line interface implementation
├── core.py             # Core functionality for document manipulation
└── functions.py        # Helper functions for text processing
```

## Screenshots

<p align="center">
  <table>
    <tr>
      <td><img src="public/project_screenshots/ss2.png" alt="" width="400"></td>
      <td><img src="public/project_screenshots/ss4.png" alt="" width="400"></td>
    </tr>
  </table>
</p>

## Developer Notes

# To Remember
python-docx can ONLY work with styles that are defined in the document.
https://stackoverflow.com/questions/51829366/bullet-lists-in-python-docx

A Run is a sequence of characters that follows the same formatting

Legal-style numbered list because python-docx don't nativally suport list
Small-caps Versalete
keeping line height and spacement between words

Existing bugs: 
1 - With we use insert after with the "The Disclosing Party ..." specifically it restarts the list
2 - Sometimes insert before/after don't keep the font styling

Contract 1 (Insert list item)
Insert this clause as section 1A, directly after the "Definitions" heading. Please match the document's style. (find the paragraph to add manually)

"Affiliate" means any entity that directly or indirectly controls, is controlled by, or is under common control with a party, where "control" means the possession, directly or indirectly, of the power to direct or cause the direction of the management and policies of such entity, whether through ownership of voting securities, by contract, or otherwise.

Contract 2 (Insert sentence between paragraph)
Insert the following between the first and second sentence in Section 11.  Please match the font of the surrounding text.

The Disclosing Party makes no representations or warranties regarding the accuracy or completeness of the Confidential Information.

Contract 3:
Insert this clause as section 11, directly after the last paragraph in section 10. If a heading is needed, format it bold and underlined, and match the document's style.

Residuals.  Nothing in this Agreement shall be construed to limit the Receiving Party's right to independently develop or acquire products or services without use of the Disclosing Party's Confidential Information, nor shall it restrict the use of any general knowledge, skills, or experience retained in unaided memory by personnel of the Receiving Party.

