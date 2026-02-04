"""
Example: Verification of New Features (Prepend & Images)

This script verifies:
1. `prepend_to_existing_doc`: Inserting content after the title.
2. `add_images_to_doc`: Inserting images with captions.
"""

import os
from utils.doc_builder import create_doc_from_template, prepend_to_existing_doc, add_images_to_doc

# 1. Create a base document
output_path = 'output/features_test.docx'
create_doc_from_template(
    template_path='templates/hebrew_template.docx',
    output_path=output_path,
    text='# כותרת המסמך\n\nטקסט רגיל במסמך.'
)
print(f"Created base document: {output_path}")

# 2. Test Prepend
prepend_text = """# הקדמה (נוסף בהתחלה)

טקסט זה נוסף באמצעות הפונקציה prepend_to_existing_doc.
הוא אמור להופיע *אחרי* הכותרת הראשית אך *לפני* הטקסט הרגיל.

- פריט מוסף 1
- פריט מוסף 2"""

prepend_to_existing_doc(output_path, prepend_text)
print("Prepended content successfully.")

# 3. Test Images
# Using the image uploaded by the user earlier for testing
# You might need to adjust this path if testing on a different machine
image_path = r'C:/Users/g_amn/.gemini/antigravity/brain/2b1858e1-fbf1-4ebb-b38e-ee550108dafe/uploaded_image_1768751506384.png'

if os.path.exists(image_path):
    add_images_to_doc(
        doc_path=output_path,
        image_paths=[image_path],
        captions=['צילום מסך לדוגמה'],
        width_inches=3.0,  # Should appear as 3 inches wide
        after_title=False   # Append at the end
    )
    print("Added image successfully.")
else:
    print(f"Warning: Test image not found at {image_path}. Skipping image test.")

print(f"\nVerification complete! Please open {output_path} to check results.")
