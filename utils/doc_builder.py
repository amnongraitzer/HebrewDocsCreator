"""
Hebrew Word Document Builder - pywin32 COM-Based RTL Document Generation

This module provides functions for creating and appending Hebrew text to Word documents
using Microsoft Word's COM interface via pywin32. This approach provides native RTL support
since Word handles all bidirectional text formatting internally.

Requirements:
    - Windows OS
    - Microsoft Word installed
    - pywin32 package (pip install pywin32)

How Other Projects Can Import and Use:
    ```python
    from utils.doc_builder import create_doc_from_template, append_to_existing_doc
    
    # Create new document
    create_doc_from_template(
        template_path='templates/hebrew_template.docx',
        output_path='output.docx',
        text='זהו טקסט בעברית'
    )
    
    # Append to existing document
    append_to_existing_doc(
        doc_path='output.docx',
        text='טקסט נוסף'
    )
    ```
"""

import os
import warnings
import unicodedata
from typing import List, Tuple

import win32com.client
import pythoncom

# Word constants
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_READING_ORDER_RTL = 0
WD_LIST_BULLET = 2
WD_STYLE_HEADING1 = -2  # Built-in Heading 1 style

# Unicode bidirectional markers (for detection only, not removed from content)
LTR_EMBED = '\u202A'  # LEFT-TO-RIGHT EMBEDDING
RTL_EMBED = '\u202B'  # RIGHT-TO-LEFT EMBEDDING
POP_DIRECTIONAL = '\u202C'  # POP DIRECTIONAL FORMATTING

# Bullet characters to detect
BULLET_CHARS = ('-', '*', '\u2022')  # hyphen, asterisk, bullet

# Right-to-Left Mark - invisible marker to anchor RTL context
RLM = '\u200F'


def _is_hebrew_char(char: str) -> bool:
    """
    Check if a single character is Hebrew.
    
    Args:
        char: Single character to check
        
    Returns:
        True if character is in Hebrew Unicode block
    """
    return 0x0590 <= ord(char) <= 0x05FF


def _contains_hebrew(text: str) -> bool:
    """
    Check if text contains Hebrew characters.
    
    Args:
        text: Input text to check
        
    Returns:
        True if text contains Hebrew characters, False otherwise
    """
    hebrew_range = range(0x0590, 0x05FF + 1)  # Hebrew Unicode block
    return any(ord(char) in hebrew_range for char in text)


def _strip_bidi_markers_for_detection(text: str) -> str:
    """
    Strip Unicode bidirectional markers ONLY for bullet detection.
    These markers are NOT removed from the actual text content.
    
    Args:
        text: Text to process
        
    Returns:
        Text with BIDI markers removed (for detection purposes only)
    """
    return text.replace(LTR_EMBED, '').replace(RTL_EMBED, '').replace(POP_DIRECTIONAL, '')


def _is_bullet_line(line: str) -> bool:
    """
    Check if a line starts with a bullet character.
    Uses Unicode-safe detection: strips BIDI markers for detection only.
    
    Args:
        line: Line to check
        
    Returns:
        True if line starts with a bullet character, False otherwise
    """
    stripped = line.lstrip()
    if not stripped:
        return False
    
    # Strip BIDI markers only for detection (not from actual content)
    detection_text = _strip_bidi_markers_for_detection(stripped)
    
    # Check if starts with bullet character
    return detection_text.startswith(BULLET_CHARS)


def _is_numbered_line(line: str) -> bool:
    """
    Check if a line starts with a number (e.g., 1., 2., 3.).
    
    Args:
        line: Line to check
        
    Returns:
        True if line starts with a number followed by period
    """
    import re
    stripped = line.lstrip()
    if not stripped:
        return False
    
    # Strip BIDI markers for detection
    detection_text = _strip_bidi_markers_for_detection(stripped)
    
    # Check if starts with number followed by period
    return bool(re.match(r'^\d+\.', detection_text))


def _is_heading_line(line: str) -> bool:
    """
    Check if a line is a heading (starts with #).
    
    Args:
        line: Line to check
        
    Returns:
        True if line starts with #
    """
    stripped = line.lstrip()
    if not stripped:
        return False
    
    # Strip BIDI markers for detection
    detection_text = _strip_bidi_markers_for_detection(stripped)
    
    # Check if starts with #
    return detection_text.startswith('#')


def _process_text(text: str) -> List[Tuple[str, str]]:
    """
    Process text into list of (content, style_type) tuples.
    Detects bullets, numbered lists, headings and splits into paragraphs.
    Adds RLM (Right-to-Left Mark) at end of lines that don't end with Hebrew.
    
    Syntax:
        # Heading text     -> Heading 1
        - bullet item      -> Bullet list
        * bullet item      -> Bullet list
        • bullet item      -> Bullet list
        1. numbered item   -> Numbered list
        2. numbered item   -> Numbered list
        regular text       -> Normal paragraph
    
    Args:
        text: Input text to process
        
    Returns:
        List of tuples: (content, style_type)
        - content: Text content (markers removed if present, RLM added if needed)
        - style_type: 'normal', 'bullet', 'number', or 'heading1'
    """
    import re
    lines = text.splitlines()
    result = []
    
    for line in lines:
        stripped = line.strip()
        
        # Empty line = paragraph break
        if not stripped:
            result.append(('', 'normal'))
            continue
        
        # Check line type in order of priority
        if _is_heading_line(line):
            # Calculate heading level (count #)
            level = len(stripped) - len(stripped.lstrip('#'))
            level = min(level, 9) # limit to Heading 9
            # Remove # marker from content
            content = stripped.lstrip('#').strip()
            style_type = f'heading{level}'
        elif _is_numbered_line(line):
            # Remove number and period from content (e.g., "1. " -> "")
            content = re.sub(r'^\d+\.\s*', '', stripped)
            style_type = 'number'
        elif _is_bullet_line(line):
            # Remove bullet marker from content
            content = stripped.lstrip('-*•\u2022').strip()
            style_type = 'bullet'
        else:
            # Regular paragraph
            content = stripped
            style_type = 'normal'
        
        # Add RLM at end if content doesn't end with Hebrew character
        # This anchors punctuation/English at the correct position in RTL
        if content and not _is_hebrew_char(content[-1]):
            content = content + RLM
        
        result.append((content, style_type))
    
    return result


def _get_word_app():
    """
    Get or create Word application instance.
    
    Returns:
        Word.Application COM object
    """
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in background
    return word


def _add_content_to_doc(doc, text: str, word_app) -> None:
    """
    Add processed content to document using Word COM Selection.
    
    Args:
        doc: Word Document COM object
        text: Text content to add
        word_app: Word Application COM object (for Selection)
    """
    processed = _process_text(text)
    
    # Use Selection for reliable text insertion
    selection = word_app.Selection
    
    # Move to end of document
    selection.EndKey(Unit=6)  # 6 = wdStory
    
    for content, style_type in processed:
        # FIRST: Remove any inherited list formatting
        selection.Range.ListFormat.RemoveNumbers()
        
        if not content:
            # Empty paragraph
            selection.TypeParagraph()
            continue
        
        # Set RTL and right alignment for the paragraph
        selection.ParagraphFormat.ReadingOrder = WD_READING_ORDER_RTL
        selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
        
        # Apply formatting based on style type
        if style_type.startswith('heading'):
            try:
                level = int(style_type.replace('heading', ''))
                # Map level to built-in style ID: H1=-2, H2=-3, ..., H6=-7
                style_id = -1 - level 
                selection.Style = doc.Styles(style_id)
            except:
                pass
        elif style_type == 'bullet':
            # Apply bullet list
            selection.Range.ListFormat.ApplyBulletDefault()
        elif style_type == 'number':
            # Apply numbered list
            selection.Range.ListFormat.ApplyNumberDefault()
        # else: normal paragraph, no special formatting
        
        # Type the content
        selection.TypeText(content)
        
        # Move to new paragraph
        selection.TypeParagraph()
        
        # Reset to Normal style after heading
        if style_type.startswith('heading'):
            selection.Style = doc.Styles("Normal")


def create_doc_from_template(template_path: str, output_path: str, text: str, title_text: str = None) -> None:
    """
    Create a new Word document from a template and add Hebrew text.
    
    Uses Word's COM interface for native RTL support. Word handles all
    bidirectional text formatting internally.
    
    Args:
        template_path: Path to .docx or .dotx template file (must exist)
        output_path: Path where output .docx file will be saved
        text: Hebrew text content to add (can contain paragraphs and bullets)
        
    Raises:
        FileNotFoundError: If template_path does not exist
        ValueError: If text encoding is invalid
        OSError: If output directory cannot be created
        
    Example:
        >>> create_doc_from_template(
        ...     template_path='templates/hebrew_template.docx',
        ...     output_path='output.docx',
        ...     text='זהו טקסט בעברית\\n- פריט רשימה'
        ... )
    """
    # Convert to absolute paths
    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(output_path)
    
    # Error handling: Check template exists
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")
    
    # Error handling: Check/create output directory
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except OSError as e:
            raise OSError(f"Cannot create output directory '{output_dir}': {e}")
    
    # Error handling: Validate encoding
    try:
        text.encode('utf-8')
    except UnicodeEncodeError as e:
        raise ValueError(f"Invalid encoding in text: {e}")
    
    # Warning: Non-Hebrew input (warn but don't fail)
    if not _contains_hebrew(text):
        warnings.warn(
            "Input text does not appear to contain Hebrew characters. "
            "Document will still be created, but RTL formatting may not be appropriate.",
            UserWarning
        )
    
    # Normalize Unicode (NFC form)
    text = unicodedata.normalize('NFC', text)
    if title_text:
        title_text = unicodedata.normalize('NFC', title_text)
    
    word = None
    doc = None
    try:
        # Get Word application
        word = _get_word_app()
        
        # Open template
        doc = word.Documents.Open(template_path)
        
        # Replace title if provided
        if title_text:
            try:
                rng = doc.Paragraphs(1).Range
                rng.Select()
                rng.Text = title_text + "\r"
                rng.Style = doc.Styles(WD_STYLE_HEADING1) 
            except Exception as e:
                print(f"Warning: Failed to set title: {e}")

        # Add content
        _add_content_to_doc(doc, text, word)
        
        # Save as new document
        doc.SaveAs(output_path, FileFormat=16)  # 16 = docx format
        
    finally:
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


def append_to_existing_doc(doc_path: str, text: str) -> None:
    """
    Append Hebrew text to an existing Word document.
    
    Uses Word's COM interface for native RTL support.
    
    Args:
        doc_path: Path to existing .docx file (must exist)
        text: Hebrew text content to append (can contain paragraphs and bullets)
        
    Raises:
        FileNotFoundError: If doc_path does not exist
        ValueError: If text encoding is invalid
        
    Example:
        >>> append_to_existing_doc(
        ...     doc_path='output.docx',
        ...     text='טקסט נוסף\\n* עוד פריט'
        ... )
    """
    # Convert to absolute path
    doc_path = os.path.abspath(doc_path)
    
    # Error handling: Check document exists
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")
    
    # Error handling: Validate encoding
    try:
        text.encode('utf-8')
    except UnicodeEncodeError as e:
        raise ValueError(f"Invalid encoding in text: {e}")
    
    # Warning: Non-Hebrew input (warn but don't fail)
    if not _contains_hebrew(text):
        warnings.warn(
            "Input text does not appear to contain Hebrew characters. "
            "Content will still be appended, but RTL formatting may not be appropriate.",
            UserWarning
        )
    
    # Normalize Unicode (NFC form)
    text = unicodedata.normalize('NFC', text)
    
    word = None
    doc = None
    try:
        # Get Word application
        word = _get_word_app()
        
        # Open existing document
        doc = word.Documents.Open(doc_path)
        
        # Add content
        _add_content_to_doc(doc, text, word)
        
        # Save document
        doc.Save()
        
    finally:
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


def prepend_to_existing_doc(doc_path: str, text: str, separator: str = "-" * 60) -> None:
    """
    Prepend Hebrew text to an existing Word document (insert after title).
    
    The content is inserted after the first heading/title paragraph,
    before any existing content. A separator line is added between
    the new content and existing content.
    
    Uses Word's COM interface for native RTL support.
    
    Args:
        doc_path: Path to existing .docx file (must exist)
        text: Hebrew text content to prepend (can contain paragraphs and bullets)
        separator: Separator string between new and existing content (default: 60 dashes)
        
    Raises:
        FileNotFoundError: If doc_path does not exist
        ValueError: If text encoding is invalid
        
    Example:
        >>> prepend_to_existing_doc(
        ...     doc_path='output.docx',
        ...     text='# כותרת חדשה\\n- פריט חדש'
        ... )
    """
    # Convert to absolute path
    doc_path = os.path.abspath(doc_path)
    
    # Error handling: Check document exists
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")
    
    # Error handling: Validate encoding
    try:
        text.encode('utf-8')
    except UnicodeEncodeError as e:
        raise ValueError(f"Invalid encoding in text: {e}")
    
    # Warning: Non-Hebrew input (warn but don't fail)
    if not _contains_hebrew(text):
        warnings.warn(
            "Input text does not appear to contain Hebrew characters. "
            "Content will still be prepended, but RTL formatting may not be appropriate.",
            UserWarning
        )
    
    # Normalize Unicode (NFC form)
    text = unicodedata.normalize('NFC', text)
    
    word = None
    doc = None
    try:
        # Get Word application
        word = _get_word_app()
        
        # Open existing document
        doc = word.Documents.Open(doc_path)
        
        # Find insertion point: after first heading/title
        # We look for the first paragraph that is a heading style
        insertion_range = None
        paragraphs = doc.Paragraphs
        
        for i in range(1, paragraphs.Count + 1):
            para = paragraphs(i)
            style_name = para.Style.NameLocal
            # Check if this is a title or heading style
            if 'Heading' in style_name or 'Title' in style_name or 'כותרת' in style_name:
                # Insert after this paragraph
                insertion_range = para.Range
                insertion_range.Collapse(Direction=0)  # 0 = wdCollapseEnd
                break
        
        # If no heading found, insert at the beginning
        if insertion_range is None:
            insertion_range = doc.Range(0, 0)
        
        # Move selection to insertion point
        insertion_range.Select()
        selection = word.Selection
        
        # Add a paragraph break first
        selection.TypeParagraph()
        
        # Process and add the new content
        processed = _process_text(text)
        
        for content, style_type in processed:
            # Remove any inherited list formatting
            selection.Range.ListFormat.RemoveNumbers()
            
            if not content:
                # Empty paragraph
                selection.TypeParagraph()
                continue
            
            # Set RTL and right alignment
            selection.ParagraphFormat.ReadingOrder = WD_READING_ORDER_RTL
            selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
            
            # Apply formatting based on style type
            if style_type.startswith('heading'):
                try:
                    level = int(style_type.replace('heading', ''))
                    # Map level to built-in style ID: H1=-2, H2=-3, ..., H6=-7
                    style_id = -1 - level 
                    selection.Style = doc.Styles(style_id)
                except:
                    pass
            elif style_type == 'bullet':
                selection.Range.ListFormat.ApplyBulletDefault()
            elif style_type == 'number':
                selection.Range.ListFormat.ApplyNumberDefault()
            
            # Type the content
            selection.TypeText(content)
            selection.TypeParagraph()
            
            # Reset to Normal style after heading
            if style_type.startswith('heading'):
                selection.Style = doc.Styles("Normal")
        
        # Add separator line
        if separator:
            selection.Range.ListFormat.RemoveNumbers()
            selection.ParagraphFormat.ReadingOrder = WD_READING_ORDER_RTL
            selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
            selection.TypeText(separator)
            selection.TypeParagraph()
        
        # Save document
        doc.Save()
        
    finally:
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


def add_images_to_doc(doc_path: str, image_paths: list, captions: list = None, 
                      width_inches: float = 2.75, after_title: bool = True, placeholder_text: str = None) -> None:
    """
    Add images to an existing Word document.
    
    Images are inserted with optional captions. By default, images are added
    after the title (first heading). Set after_title=False to append at end.
    If placeholder_text is provided, it searches for that text and replaces it with images.
    
    Uses Word's COM interface for native image handling.
    
    Args:
        doc_path: Path to existing .docx file (must exist)
        image_paths: List of image file paths to add
        captions: Optional list of caption strings (same length as image_paths)
        width_inches: Width of images in inches (default: 2.75)
        after_title: If True, insert after first heading; if False, append at end
        
    Raises:
        FileNotFoundError: If doc_path or any image_path does not exist
        ValueError: If captions length doesn't match image_paths length
        
    Example:
        >>> add_images_to_doc(
        ...     doc_path='output.docx',
        ...     image_paths=['image1.jpg', 'image2.jpg'],
        ...     captions=['תמונה 1', 'תמונה 2']
        ... )
    """
    # Convert to absolute path
    doc_path = os.path.abspath(doc_path)
    
    # Error handling: Check document exists
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")
    
    # Check all images exist
    for img_path in image_paths:
        abs_img_path = os.path.abspath(img_path)
        if not os.path.exists(abs_img_path):
            raise FileNotFoundError(f"Image not found: {abs_img_path}")
    
    # Validate captions length
    if captions is not None and len(captions) != len(image_paths):
        raise ValueError(f"Captions length ({len(captions)}) must match image_paths length ({len(image_paths)})")
    
    if not image_paths:
        return  # Nothing to do
    
    word = None
    doc = None
    try:
        # Get Word application
        word = _get_word_app()
        
        # Open existing document
        doc = word.Documents.Open(doc_path)
        
        # Find insertion point
        if placeholder_text:
            # Find and select placeholder
            selection = word.Selection
            selection.Find.ClearFormatting()
            found = selection.Find.Execute(FindText=placeholder_text)
            
            if found:
                # Delete placeholder text
                selection.Delete()
                insertion_range = selection.Range
            else:
                # Fallback to end if not found
                insertion_range = doc.Range()
                insertion_range.Collapse(Direction=0) 
                
        elif after_title:
            # Find position after first heading
            insertion_range = None
            paragraphs = doc.Paragraphs
            
            for i in range(1, paragraphs.Count + 1):
                para = paragraphs(i)
                style_name = para.Style.NameLocal
                if 'Heading' in style_name or 'Title' in style_name or 'כותרת' in style_name:
                    insertion_range = para.Range
                    insertion_range.Collapse(Direction=0)  # wdCollapseEnd
                    break
            
            if insertion_range is None:
                insertion_range = doc.Range(0, 0)
        else:
            # Append at end
            insertion_range = doc.Range()
            insertion_range.Collapse(Direction=0)  # wdCollapseEnd
        
        # Move selection to insertion point
        insertion_range.Select()
        selection = word.Selection
        
        # Add header for images section
        selection.TypeParagraph()
        selection.ParagraphFormat.ReadingOrder = WD_READING_ORDER_RTL
        selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
        selection.Font.Bold = True
        selection.TypeText("עזרים ויזואליים:")
        selection.Font.Bold = False
        selection.TypeParagraph()
        
        # Convert width to points (1 inch = 72 points)
        width_points = width_inches * 72
        
        # Add each image
        for idx, img_path in enumerate(image_paths):
            abs_img_path = os.path.abspath(img_path)
            
            # Center alignment for image
            selection.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
            
            # Add the image
            inline_shape = selection.InlineShapes.AddPicture(
                FileName=abs_img_path,
                LinkToFile=False,
                SaveWithDocument=True
            )
            
            # Ensure aspect ratio is locked before setting width
            inline_shape.LockAspectRatio = -1  # msoTrue
            
            # Set width (height scales proportionally)
            inline_shape.Width = width_points
            
            selection.TypeParagraph()
            
            # Add caption if provided
            if captions and captions[idx]:
                selection.ParagraphFormat.Alignment = 1  # Center
                selection.Font.Size = 8
                selection.Font.Color = 6710886  # Gray color (RGB: 102, 102, 102)
                selection.TypeText(captions[idx])
                selection.Font.Size = 12  # Reset to default
                selection.Font.Color = 0  # Reset to black
                selection.TypeParagraph()
        
        # Save document
        doc.Save()
        
    finally:
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()
