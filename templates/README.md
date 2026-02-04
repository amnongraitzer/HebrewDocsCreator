# Templates Directory

Place your Hebrew RTL template file (`hebrew_template.dotx`) in this directory.

## Template Requirements

The template must be pre-configured in Microsoft Word with:
- Normal style: RTL, right alignment, Hebrew font, language=he-IL
- List Bullet style: RTL, right alignment, Hebrew font, bullets on right
- Heading styles (1-3): RTL, right alignment, Hebrew font
- Document defaults: Hebrew language, RTL direction

## Creating the Template

See [../TEMPLATE_GUIDE.md](../TEMPLATE_GUIDE.md) for detailed step-by-step instructions.

## Quick Start

1. Follow the instructions in TEMPLATE_GUIDE.md
2. Save your template as `hebrew_template.dotx` in this directory
3. Use the template in your code:

```python
from utils.doc_builder import create_doc_from_template

create_doc_from_template(
    template_path='templates/hebrew_template.dotx',
    output_path='output.docx',
    text='טקסט בעברית'
)
```
