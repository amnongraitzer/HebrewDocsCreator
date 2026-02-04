"""
Example: Creating document with Heading 1, bullet list, and numbered list

This example demonstrates how to create a document with various formatting:
- Headings (using # prefix)
- Bullet lists (using -, *, or • prefix)
- Numbered lists (using 1., 2., etc. prefix)
"""

from utils.doc_builder import create_doc_from_template

# Hebrew text with heading, bullet list, and numbered list
text_with_all_styles = """# כותרת ראשית של המסמך

זהו פסקה רגילה שמסבירה את תוכן המסמך.

רשימה עם נקודות:

- פריט ראשון ברשימה
- פריט שני ברשימה
- פריט שלישי ברשימה

רשימה ממוספרת:

1. שלב ראשון בתהליך
2. שלב שני בתהליך
3. שלב שלישי בתהליך

# כותרת משנית

עוד טקסט רגיל אחרי הכותרת."""

# Create document
create_doc_from_template(
    template_path='templates/hebrew_template.docx',
    output_path='output/all_styles.docx',
    text=text_with_all_styles
)

print("Document with all styles created: output/all_styles.docx")
print("Includes: Heading 1, bullet list, and numbered list")
