"""
Example: Appending Heading 1, bullet list, and numbered list to existing document

This example demonstrates how to append various formatted content to an existing document.
"""

from utils.doc_builder import append_to_existing_doc

# Hebrew text with heading, bullet list, and numbered list to append
text_to_append = """# כותרת נוספת שנוספה

טקסט נוסף שנוסף למסמך הקיים.

פריטים נוספים:

- פריט א
- פריט ב
- פריט ג

משימות חדשות:

1. משימה ראשונה
2. משימה שנייה
3. משימה שלישית"""

# Append to existing document
append_to_existing_doc(
    doc_path='output/all_styles.docx',
    text=text_to_append
)

print("Content appended successfully to: output/all_styles.docx")
print("Appended: Heading 1, bullet list, and numbered list")
