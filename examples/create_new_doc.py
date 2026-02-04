"""
Example: Creating a new document from template

This example demonstrates how to create a new Word document from a Hebrew template
and add Hebrew text content.
"""

from utils.doc_builder import create_doc_from_template

# Hebrew text content
hebrew_text = """זהו מסמך בעברית שנוצר מתבנית.

המסמך מכיל מספר פסקאות בעברית עם פיסוק נכון.
כל הפסקאות מיושרות לימין ומשתמשות בסגנון מהתבנית.

זהו פסקה נוספת שמדגימה את השימוש בתבנית."""

# Create new document from template
create_doc_from_template(
    template_path='templates/hebrew_template.docx',
    output_path='output/new_document.docx',
    text=hebrew_text
)

print("Document created successfully: output/new_document.docx")
