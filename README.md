# Hebrew Word Document Builder

A reusable Python module for generating Word (.docx) documents in Hebrew using a predefined .dotx template. All RTL formatting is controlled by the template, ensuring consistent and reliable Hebrew document generation.

## Architecture Overview

This module follows a **template-based approach** where all RTL formatting (alignment, direction, fonts, language) comes from a pre-configured Word template (.dotx file). The code is minimal and style-agnostic, focusing only on content insertion.

### Key Principles

- **Template controls all formatting**: RTL direction, alignment, fonts, language settings
- **Code only inserts content**: No RTL flags, no alignment overrides (except minimal safety check when appending)
- **Reusable across projects**: Simple function-based API
- **No hard-coded styling**: All formatting comes from template

### Benefits

1. **Simpler Code**: ~150 lines instead of ~300+ lines with XML workarounds
2. **More Reliable**: Word's native template configuration avoids python-docx bugs
3. **Better Maintainability**: Style changes happen in Word template, not Python code
4. **Consistent Results**: All documents using same template have identical formatting
5. **Easier Debugging**: Can test template in Word directly before using in code

## Installation

```bash
pip install -r requirements.txt
```

## Template Requirements

**IMPORTANT**: You must create a properly configured Hebrew RTL template (.dotx file) before using this module.

The template must be pre-configured in Word with:
- **Normal style**: RTL paragraph direction, Hebrew font, right alignment, language=he-IL
- **List Bullet style**: RTL, Hebrew font, right alignment, bullets on right
- **Heading styles** (Heading 1, 2, 3): RTL, Hebrew font, right alignment
- **Document defaults**: Hebrew language, RTL direction

See [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) for detailed step-by-step instructions.

## API Reference

### `create_doc_from_template(template_path, output_path, text)`

Create a new Word document from a template and add Hebrew text.

**Parameters:**
- `template_path` (str): Path to .dotx template file (must exist)
- `output_path` (str): Path where output .docx file will be saved
- `text` (str): Hebrew text content to add (can contain paragraphs and bullets)

**Raises:**
- `FileNotFoundError`: If template_path does not exist
- `ValueError`: If text encoding is invalid
- `OSError`: If output directory cannot be created

**Example:**
```python
from utils.doc_builder import create_doc_from_template

create_doc_from_template(
    template_path='templates/hebrew_template.dotx',
    output_path='output/document.docx',
    text='זהו טקסט בעברית'
)
```

### `append_to_existing_doc(doc_path, text)`

Append Hebrew text to an existing Word document.

**Parameters:**
- `doc_path` (str): Path to existing .docx file (must exist)
- `text` (str): Hebrew text content to append (can contain paragraphs and bullets)

**Raises:**
- `FileNotFoundError`: If doc_path does not exist
- `ValueError`: If text encoding is invalid

**Note:** Includes minimal safety check: if new paragraph alignment is not RIGHT, sets it to RIGHT. This is the ONLY exception to the "no alignment override" rule.

**Example:**
```python
from utils.doc_builder import append_to_existing_doc

append_to_existing_doc(
    doc_path='output/document.docx',
    text='טקסט נוסף'
)
```

## Usage Examples

### Creating a New Document

```python
from utils.doc_builder import create_doc_from_template

hebrew_text = """זהו מסמך בעברית שנוצר מתבנית.

המסמך מכיל מספר פסקאות בעברית עם פיסוק נכון.
כל הפסקאות מיושרות לימין ומשתמשות בסגנון מהתבנית."""

create_doc_from_template(
    template_path='templates/hebrew_template.dotx',
    output_path='output/new_document.docx',
    text=hebrew_text
)
```

### Appending to Existing Document

```python
from utils.doc_builder import append_to_existing_doc

additional_text = """זהו תוכן נוסף שנוסף למסמך הקיים.

התוכן החדש יתווסף בסוף המסמך."""

append_to_existing_doc(
    doc_path='output/new_document.docx',
    text=additional_text
)
```

### Creating Bulleted Lists

The module automatically detects bullet markers (`-`, `*`, `•`) and applies the template's List Bullet style:

```python
from utils.doc_builder import create_doc_from_template

text_with_bullets = """רשימת קניות:

- חלב
- לחם
- ביצים

רשימה נוספת:

* פריט ראשון
* פריט שני"""

create_doc_from_template(
    template_path='templates/hebrew_template.dotx',
    output_path='output/bulleted_list.docx',
    text=text_with_bullets
)
```

See the [examples/](examples/) folder for more complete examples.

## How It Works

1. **Load Template**: Opens the .dotx template file
2. **Process Text**: Splits text into paragraphs, detects bullets
3. **Apply Styles**: Uses template styles (Normal, List Bullet, etc.)
4. **Add Content**: Inserts text content (preserves Unicode, including BIDI markers)
5. **Save Document**: Saves the final .docx file

### Bullet Detection

- Detects lines starting with: `-`, `*`, or `•` (\u2022)
- Strips Unicode bidirectional markers (U+202A, U+202B, U+202C) **only for detection**
- **Preserves BIDI markers in actual text content**
- Removes bullet markers from content before inserting

### Text Processing

- Handles multiple paragraphs (split by newlines)
- Empty lines create paragraph breaks
- Unicode normalization (NFC form)
- UTF-8 encoding validation

## Importing in Other Projects

```python
# Option 1: Import from utils package
from utils.doc_builder import create_doc_from_template, append_to_existing_doc

# Option 2: Import from utils package (shorter)
from utils import create_doc_from_template, append_to_existing_doc

# Use the functions
create_doc_from_template(
    template_path='path/to/template.dotx',
    output_path='output.docx',
    text='טקסט בעברית'
)
```

## Do/Don't Rules

### DO:
- ✅ Use template styles (Normal, List Bullet, Heading 1-3)
- ✅ Let template control all formatting
- ✅ Create template with proper Hebrew RTL configuration
- ✅ Use bullet markers (`-`, `*`, `•`) for lists

### DON'T:
- ❌ Use `run.font.rtl = True` (Word handles bidi internally)
- ❌ Override alignment (except minimal safety check when appending)
- ❌ Modify or create styles programmatically
- ❌ Hard-code any styling logic
- ❌ Use XML manipulation for formatting

## Error Handling

The module includes comprehensive error handling:

- **Missing template**: Raises `FileNotFoundError` with clear message
- **Missing output directory**: Creates directory if possible, raises `OSError` if not
- **Invalid encoding**: Raises `ValueError` with details
- **Non-Hebrew input**: Issues warning but does not fail (document still created)

## Testing

Run the test suite:

```bash
python -m pytest tests/
# or
python -m unittest tests.test_doc_builder
```

Tests verify:
- Paragraph alignment remains RTL
- Template styles are preserved
- Hebrew text is readable and correctly ordered
- Bullet lists appear correctly
- Error handling works correctly

## Known Limitations

### Template Dependency
- **Template is REQUIRED**: Must create template with proper Hebrew RTL configuration
- Template must be correctly configured in Word before use
- See [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) for instructions

### Word/OS Requirements
- **Word Editing Language**: User's Word must have Hebrew added as editing/authoring language
- **Hebrew Proofing Tools**: Required for spell-check/grammar (text still renders without them)
- **Font Installation**: If specified Hebrew font not installed, Word will substitute

### Font Fallback
- If template uses a font not installed on user's system, Word will substitute
- Recommend using common fonts (Arial, Times New Roman) that support Hebrew

### Advanced RTL Features
- Some paragraph direction features beyond alignment depend on Word version and user settings
- Template must be properly configured for best results

## File Structure

```
utils/
    doc_builder.py          # Main module with public functions
    __init__.py             # Package initialization
templates/
    hebrew_template.dotx    # Pre-configured Hebrew RTL template
examples/
    create_new_doc.py       # Example: Creating new document
    append_to_existing.py   # Example: Appending to existing doc
    bulleted_list.py        # Example: Bulleted lists
tests/
    test_doc_builder.py     # Automated tests
README.md                   # This file
TEMPLATE_GUIDE.md          # Template creation guide
requirements.txt           # Dependencies
```

## License

This module is provided as-is for use in Hebrew document generation projects.

## Support

For issues or questions:
1. Ensure template is correctly configured (see TEMPLATE_GUIDE.md)
2. Check that Word has Hebrew editing language enabled
3. Verify template path is correct
4. Review error messages for specific issues
