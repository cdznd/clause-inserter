import argparse
from typing import Optional
from core import run_inserter
from cli import display_document_paragraphs, get_list_style, get_list_definition_type, get_position, get_sentence_index

def clause_inserter(
    document_path: str,
    paragraph_number: int,
    position: str,
    follow_list_style: bool,
    clause_text: str,
    sentence_index: Optional[int] = None,
    list_definition_type: str = "hardcoded",
):

    run_inserter(
        document_path=document_path,
        paragraph_number=paragraph_number,
        position=position,
        follow_list_style=follow_list_style,
        clause_text=clause_text,
        sentence_index=sentence_index,
        list_definition_type=list_definition_type
    )

    return True

def main():

    """Main entry point for the CLI tool."""
    parser = argparse.ArgumentParser(description="Insert legal clauses into Word documents with precise formatting control")

    # Doc path required positional argument
    parser.add_argument("document_path", type=str, help="Path to the Word document")

    args = parser.parse_args()
    
    # Interactive mode - select paragraph
    paragraph_number = display_document_paragraphs(args.document_path) + 1  # Adding 1 as we'll subtract it later
    
    # Interactive mode - select position
    position = get_position()
    
    # For "between" position, legal style is not applicable
    follow_list_style = False
    list_definition_type = None
    if position != "between":
        # Interactive mode - select legal style
        follow_list_style = get_list_style()

        # If legal style is selected, prompt for list definition type
        if follow_list_style and list_definition_type is None:
            list_definition_type = get_list_definition_type()

    # Interactive mode - if position is between, select sentence index
    sentence_index = None
    if position == "between" and sentence_index is None:
        sentence_index = get_sentence_index(args.document_path, paragraph_number - 1)  # -1 because we use 0-based indexing

    # Get clause text if not provided
    clause_text = None
    if clause_text is None:
        clause_text = input("\nEnter the clause text to insert: ")

    # Check if we have everything we need
    if position == "between" and sentence_index is None:
        print("Error: Sentence index is required when position is 'between'")
        return 1

    success = clause_inserter(
        document_path=args.document_path,
        paragraph_number=paragraph_number - 1,  # subtracting 1 considering we start counting from 1
        position=position,
        follow_list_style=follow_list_style,
        clause_text=clause_text,
        sentence_index=sentence_index,
        list_definition_type=list_definition_type
    )

    if success:
        return 0
    else:
        return 1

if __name__ == "__main__":
    exit(main()) 