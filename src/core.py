import os
from docx import Document
from typing import Optional
from functions import insert_list_item_after, insert_paragraph_after, insert_paragraph_before, insert_list_item_before, insert_text_between_sentences, insert_manual_numbered_item_after

def run_inserter(
    document_path: str,
    paragraph_number: int,
    position: str,
    follow_list_style: bool,
    clause_text: str,
    sentence_index: Optional[int] = None,
    list_definition_type: str = "hardcoded",
):

    document = Document(document_path)

    selected_paragraph = document.paragraphs[paragraph_number]

    if position == 'after':
        if follow_list_style:
            if list_definition_type == "defined":
                insert_list_item_after(document, selected_paragraph, clause_text)
            else:  # hardcoded
                insert_manual_numbered_item_after(selected_paragraph, clause_text)
        else:
            insert_paragraph_after(selected_paragraph, clause_text, selected_paragraph.style)
    elif position == 'before':
        if follow_list_style:
            if list_definition_type == "defined":
                insert_list_item_before(document, selected_paragraph, clause_text)
            else:  # hardcoded
                # insert_manual_numbered_item_before(document, selected_paragraph, clause_text)
                print('')
        else:
            insert_paragraph_before(selected_paragraph, clause_text, selected_paragraph.style)
    elif position == 'between':
        if sentence_index is None:
            raise ValueError("sentence_index is required for 'between' position")
        insert_text_between_sentences(selected_paragraph, clause_text, sentence_index)
    else:
        raise ValueError("Position must be 'after', 'before', or 'between'")

    output_folder_path = './processed_documents'

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Save
    document.save(f'{output_folder_path}/new_file.docx')
