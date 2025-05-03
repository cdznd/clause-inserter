from docx import Document
import textwrap

def display_document_paragraphs(document_path: str):
    """
    Displays all paragraphs in the document with their indices and lets the user select one.
    
    Args:
        document_path: Path to the Word document
        
    Returns:
        int: The selected paragraph index (0-based)
    """
    document = Document(document_path)
    
    print("\nDocument Paragraphs:")
    print("-" * 80)
    
    for i, para in enumerate(document.paragraphs):
        # Skip empty paragraphs
        if not para.text.strip():
            continue
            
        # Truncate and format text for display
        preview = textwrap.shorten(para.text.strip(), width=70, placeholder="...")
        print(f"[{i}] {preview}")
    
    print("-" * 80)
    
    while True:
        try:
            selection = int(input("Enter the paragraph number to target: "))
            if 0 <= selection < len(document.paragraphs):
                return selection
            else:
                print(f"Please enter a number between 0 and {len(document.paragraphs) - 1}")
        except ValueError:
            print("Please enter a valid number")

def get_list_style():
    """
    Prompt the user to select whether to use legal style formatting.
    
    Returns:
        bool: True for legal style, False otherwise
    """
    print("\nDo you want to follow list style formatting?")
    print("[1] Yes")
    print("[2] No")
    
    while True:
        try:
            choice = int(input("Enter your choice (1-2): "))
            if choice == 1:
                return True
            elif choice == 2:
                return False
            else:
                print("Please enter either 1 or 2")
        except ValueError:
            print("Please enter a valid number")

def get_list_definition_type():
    """
    Prompt the user to select the type of list definition to use.
    
    Returns:
        str: The list definition type ('defined' or 'hardcoded')
    """
    print("\nPlease select the list definition type:")
    print("[1] Use defined list formatting (from document)")
    print("[2] Use hardcoded list formatting")
    
    while True:
        try:
            choice = int(input("Enter your choice (1-2): "))
            if choice == 1:
                return "defined"
            elif choice == 2:
                return "hardcoded"
            else:
                print("Please enter either 1 or 2")
        except ValueError:
            print("Please enter a valid number")

def get_position():
    """
    Prompt the user to select the insertion position.
    
    Returns:
        str: The position (before, after, or between)
    """
    print("\nWhere would you like to insert the clause?")
    print("[1] Before paragraph")
    print("[2] After paragraph")
    print("[3] Between sentences in paragraph")
    
    while True:
        try:
            choice = int(input("Enter your choice (1-3): "))
            if choice == 1:
                return "before"
            elif choice == 2:
                return "after"
            elif choice == 3:
                return "between"
            else:
                print("Please enter a number between 1 and 3")
        except ValueError:
            print("Please enter a valid number")

def get_sentence_index(document_path: str, paragraph_index: int):
    """
    Display the selected paragraph split into sentences and prompt for sentence index.
    
    Args:
        document_path: Path to the Word document
        paragraph_index: Index of the paragraph
        
    Returns:
        int: The selected sentence index
    """
    document = Document(document_path)
    paragraph = document.paragraphs[paragraph_index]
    
    # Split text into sentences (simple split by period + space)
    sentences = paragraph.text.split('. ')
    
    print("\nParagraph sentences:")
    print("-" * 80)
    
    for i, sentence in enumerate(sentences):
        print(f"[{i}] {sentence}")
    
    print("-" * 80)
    
    while True:
        try:
            selection = int(input("Enter the sentence index to insert after: "))
            if 0 <= selection < len(sentences):
                return selection
            else:
                print(f"Please enter a number between 0 and {len(sentences) - 1}")
        except ValueError:
            print("Please enter a valid number")