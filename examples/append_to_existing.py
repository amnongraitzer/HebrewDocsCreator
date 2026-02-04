"""
Example: Adding Hebrew text to existing .docx

This example demonstrates how to append Hebrew text content to an existing
Word document. The document must already exist.
"""

from utils.doc_builder import append_to_existing_doc

# Hebrew text to append
additional_text = """זהו תוכן נוסף שנוסף למסמך הקיים.

התוכן החדש יתווסף בסוף המסמך וישתמש בסגנונות מהתבנית.
אם המסמך לא נוצר מהתבנית, עדיין יתווסף תוכן עם יישור לימין."""

# Append to existing document
append_to_existing_doc(
    doc_path='output/new_document.docx',
    text=additional_text
)

print("Content appended successfully to: output/new_document.docx")
