# Clause Inserter
A Python CLI tool for programmatically inserting clauses into legal MS Word documents (.docx) with accurate formatting control, built using the [python-docx](https://python-docx.readthedocs.io/en/latest/) library.

## 📌 Overview
This CLI tool allows users to insert clauses, sentences, or paragraphs into **MS Word documents** while maintaining consistent formatting, numbering, and styling. Built using the **python-docx** library, it features an interactive CLI interface. The tool was designed to support a wide range of document types written in many different styles, ensuring flexibility across diverse formatting conventions.

**Example Use Cases**

**Contract 1 – Insert List Item**

Insert the following as section 1A, directly after the "Definitions" heading. Match the style and numbering of the document:

> “Affiliate” means any entity that directly or indirectly controls, is controlled by, or is under common control with a party, where “control” means the possession, directly or indirectly, of the power to direct or cause the direction of the management and policies of such entity, whether through ownership of voting securities, by contract, or otherwise.

**Contract 2 – Insert Sentence Within Paragraph**

Insert the sentence below between the first and second sentence in Section 11. Match the surrounding font:

> The Disclosing Party makes no representations or warranties regarding the accuracy or completeness of the Confidential Information.

**Contract 3 – Insert New Clause with Heading**

Insert the clause below as Section 11, directly after the last paragraph in Section 10. Add a heading in bold and underlined if needed:

> Residuals. Nothing in this Agreement shall be construed to limit the Receiving Party’s right to independently develop or acquire products or services without use of the Disclosing Party’s Confidential Information, nor shall it restrict the use of any general knowledge, skills, or experience retained in unaided memory by personnel of the Receiving Party.

## 🧩 Features
- CLI Interactive interface
- Insert clauses before or after paragraphs
- Insert text between sentences in a paragraph
- Maintain document numbering and list styles
- Preserve formatting of surrounding text
- Interactive paragraph selection

## ✅ Getting Started

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

### Usage

Run the tool from the command line by providing a Word document path:

```bash
python src/clause_inserter.py path/to/document.docx
```

The tool will guide you through an interactive process to:
1. Select a target paragraph
2. Choose insertion position (before, after, or between sentences)
3. Configure formatting options
4. Insert your clause text

## 💻 Technologies
- **Python**: Core programming language
- **python-docx**: Library for working with Word documents

## 🏗️ Project Structure
```
src/
├── clause_inserter.py  # Main application file
├── cli.py              # Command line interface implementation
├── core.py             # Core functionality for document manipulation
└── functions.py        # Helper functions for text processing
```
