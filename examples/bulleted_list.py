"""
Example: Creating bulleted list using template RTL list style

This example demonstrates how to create a document with bulleted lists.
The module automatically detects bullet markers (-, *, •) and applies
the template's List Bullet style.
"""

from utils.doc_builder import create_doc_from_template

# Hebrew text with bulleted lists
text_with_bullets = """רשימת קניות:

- חלב.
- לחם.
- ביצים!
- גבינה@

רשימה נוספת עם כוכבית:

* פריט ראשון ברשימה with english at the end
* Second item שממשיך באנגלית
* פריט שלישי רגיל ברשימה

רשימה עם נקודה:

• משימה אחת
• משימה שנייה
• משימה שלישית"""

# Create document with bulleted lists
create_doc_from_template(
    template_path='templates/hebrew_template.docx',
    output_path='output/bulleted_list.docx',
    text=text_with_bullets
)

print("Document with bulleted lists created: output/bulleted_list.docx")
print("Bullets are automatically detected and formatted using template's List Bullet style")
