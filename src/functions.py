import re
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from haggis.files.docx import list_number

def insert_paragraph_after(paragraph, text, style=None):
    new_p = OxmlElement('w:p')
    paragraph._element.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    new_paragraph.add_run(text)
    if style:
        new_paragraph.style = style

def insert_paragraph_before(paragraph, text, style=None):
    new_p = OxmlElement('w:p')
    paragraph._element.addprevious(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    new_paragraph.add_run(text)
    if style:
        new_paragraph.style = style

def insert_list_item_after(document, paragraph, text):
    new_p = OxmlElement('w:p')
    paragraph._element.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if text:
        words = text.split(maxsplit=1)
        first_word = words[0]
        rest = words[1] if len(words) > 1 else ""
        run1 = new_paragraph.add_run(first_word)
        run1.bold = True
        if rest:
            new_paragraph.add_run(' ' + rest)
    list_number(
        doc=document,
        par=new_paragraph,
        prev=paragraph,
    )

def insert_list_item_before(document, paragraph, text):
    new_p = OxmlElement('w:p')
    paragraph._element.addprevious(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if text:
        words = text.split(maxsplit=1)
        first_word = words[0]
        rest = words[1] if len(words) > 1 else ""
        run1 = new_paragraph.add_run(first_word)
        run1.bold = True
        if rest:
            new_paragraph.add_run(' ' + rest)
    list_number(
        doc=document,
        par=new_paragraph,
        prev=paragraph,
    )

# bug: If the insert text has a dot at the end it'll duplicate
def insert_text_between_sentences(paragraph, insert_text, sentence_index):

    # Step 1: Save all runs and their texts / a run is a sequence of characters that follow the same formatting
    run_infos = [(run, run.text) for run in paragraph.runs]

    # Full paragraph text from runs
    full_text = ''.join(text for run, text in run_infos)

    # Split into sentences
    sentences = full_text.split('. ')

    if sentence_index < 0 or sentence_index >= len(sentences):
        raise IndexError("sentence_index out of range.")

    # Insert new text
    insert_text = insert_text.strip()
    sentences.insert(sentence_index + 1, insert_text)

    # Rebuild full text
    new_full_text = '. '.join(sentences)
    if paragraph.text.endswith('.') and not new_full_text.endswith('.'):
        new_full_text += '.'


    # remove all original runs
    p_element = paragraph._element # _element to access the raw XML from the element
    for run in paragraph.runs:
        p_element.remove(run._element)

    # Step 3: Recreate runs by cloning original run XML
    idx = 0
    for old_run, old_text in run_infos:
        if idx >= len(new_full_text):
            break
        
        # Size of old text
        take_len = len(old_text)
        run_text = new_full_text[idx:idx + take_len] # Extracting the portion of the new_full_text

        if not run_text:
            continue

        # Clone the original run XML
        new_run_element = deepcopy(old_run._element) # deep copy for copying all entries
        new_run_element.text = run_text

        # Append to paragraph
        p_element.append(new_run_element)

        idx += take_len

    # Handle leftover text
    if idx < len(new_full_text):
        last_run = run_infos[-1][0]
        leftover_run_element = deepcopy(last_run._element)
        leftover_run_element.text = new_full_text[idx:]
        p_element.append(leftover_run_element)


def insert_manual_numbered_item_after(paragraph, text):

    # Detect current number using regex
    match = re.match(r'(\d+)\.\s', paragraph.text)
    if not match:
        raise ValueError("Paragraph does not start with a number followed by a dot and space.")

    current_number = int(match.group(1))
    next_number = current_number + 1

    # Create the new paragraph
    new_p = OxmlElement('w:p')
    paragraph._element.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)

    new_paragraph.style = paragraph.style
    if paragraph.paragraph_format:
        new_paragraph.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
        new_paragraph.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
        new_paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_before
        new_paragraph.paragraph_format.space_after = paragraph.paragraph_format.space_after
        # for missing tab
        new_paragraph.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
        # line height
        new_paragraph.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
        new_paragraph.paragraph_format.line_spacing_rule = paragraph.paragraph_format.line_spacing_rule

    # Grab font and size from the first run
    font_name = None
    font_size = None
    if paragraph.runs:
        first_run = paragraph.runs[0]
        font_name = first_run.font.name
        font_size = first_run.font.size

    # for later turn the first word bold and underline
    if text:
        parts = text.strip().split(' ', 1)
        first_word = parts[0]
        rest_text = parts[1] if len(parts) > 1 else ''
    else:
        first_word = ''
        rest_text = ''

    # First run: number and dot (normal)
    number_run = new_paragraph.add_run(f"{next_number}. ")
    if font_name:
        number_run.font.name = font_name
    if font_size:
        number_run.font.size = font_size 
    number_run.bold = True  # Make number bold

    # Second run: first word (bold + underline)
    first_word_run = new_paragraph.add_run(first_word)
    if font_name:
        first_word_run.font.name = font_name
    if font_size:
        first_word_run.font.size = font_size
    first_word_run.bold = True
    first_word_run.underline = True

    # Third run: space + rest of text (normal)
    if rest_text:
        rest_run = new_paragraph.add_run(f" {rest_text}")
        if font_name:
            rest_run.font.name = font_name
        if font_size:
            rest_run.font.size = font_size

    # update the following paragraphs
    following_para = new_paragraph._p.getnext()
    while following_para is not None:
        next_paragraph = Paragraph(following_para, paragraph._parent)
        original_text = next_paragraph.text

        next_match = re.match(r'(\d+)\.\s', original_text)
        if not next_match:
            break  # if not a numbered item, stop

        old_number = int(next_match.group(1))
        updated_number = old_number + 1

        if next_paragraph.runs:
            first_run = next_paragraph.runs[0]
            original_text = first_run.text
            new_first_run_text = re.sub(r'^\d+\.\s', f"{updated_number}. ", original_text, count=1)
            first_run.text = new_first_run_text
        else:
            next_paragraph.add_run(f"{updated_number}. ")

        following_para = following_para.getnext()