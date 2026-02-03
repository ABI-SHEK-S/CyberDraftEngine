from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

def replace_placeholder_in_paragraph(paragraph, replacements):
    """
    Replace placeholders in a paragraph while preserving run-level formatting.
    Handles cases where placeholders span multiple runs.
    
    Args:
        paragraph: The docx paragraph object to process.
        replacements: A dictionary mapping placeholders to their replacement values.
    Returns:
        bool: True if any replacements were made, False otherwise.
    """
    # Collect all runs and their text
    runs = paragraph.runs
    if not runs:
        return False

    # Combine text from all runs to search for placeholders
    full_text = ''.join(run.text for run in runs)
    original_text = full_text
    modified = False

    # Perform replacements on the full text
    for placeholder, value in replacements.items():
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, value)
            modified = True

    if not modified:
        return False

    # Clear all runs in the paragraph
    paragraph.clear()

    # If the paragraph text was modified, distribute the new text across runs
    # while preserving original formatting
    current_pos = 0
    for run in runs:
        if not run.text:
            continue  # Skip empty runs
        # Get the portion of the new text that corresponds to this run's length
        run_length = len(run.text)
        new_run_text = full_text[current_pos:current_pos + run_length]
        if new_run_text:
            new_run = paragraph.add_run(new_run_text)
            # Copy formatting from the original run
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.name = run.font.name
            new_run.font.size = run.font.size
            new_run.font.color.rgb = run.font.color.rgb if run.font.color else None
            # Handle East Asian fonts if needed
            if run.font.name:
                new_run._element.rPr.rFonts.set(qn('w:eastAsia'), run.font.name)
        current_pos += run_length

    # Add any remaining text as a new run with the last run's formatting
    if current_pos < len(full_text):
        remaining_text = full_text[current_pos:]
        last_run = runs[-1] if runs else None
        new_run = paragraph.add_run(remaining_text)
        if last_run:
            new_run.bold = last_run.bold
            new_run.italic = last_run.italic
            new_run.underline = last_run.underline
            new_run.font.name = last_run.font.name
            new_run.font.size = last_run.font.size
            new_run.font.color.rgb = last_run.font.color.rgb if last_run.font.color else None
            if last_run.font.name:
                new_run._element.rPr.rFonts.set(qn('w:eastAsia'), last_run.font.name)

    # Clean up any empty runs
    for run in paragraph.runs[:]:
        if not run.text.strip():
            run._element.getparent().remove(run._element)

    return modified