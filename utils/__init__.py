"""
Hebrew Word Document Builder - Reusable Module

This package provides functions for creating and appending Hebrew text to Word documents
using a predefined .dotx template with RTL formatting.

Public API:
    - create_doc_from_template: Create a new document from template
    - append_to_existing_doc: Append content to an existing document
    - prepend_to_existing_doc: Prepend content (insert after title)
    - add_images_to_doc: Add images with captions to a document

Example:
    from utils.doc_builder import create_doc_from_template, prepend_to_existing_doc
    
    # Create new document
    create_doc_from_template(
        template_path='templates/hebrew_template.dotx',
        output_path='output.docx',
        text='זהו טקסט בעברית'
    )
    
    # Prepend to existing document (insert at top after title)
    prepend_to_existing_doc(
        doc_path='output.docx',
        text='טקסט חדש בראש המסמך'
    )
"""

from utils.doc_builder import (
    create_doc_from_template,
    append_to_existing_doc,
    prepend_to_existing_doc,
    add_images_to_doc,
)

__all__ = [
    'create_doc_from_template',
    'append_to_existing_doc',
    'prepend_to_existing_doc',
    'add_images_to_doc',
]
