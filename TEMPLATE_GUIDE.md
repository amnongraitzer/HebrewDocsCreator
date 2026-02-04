# Hebrew Template Creation Guide

This guide provides step-by-step instructions for creating a Hebrew RTL Word template (.dotx file) that works with the doc_builder module.

## Overview

The template must be pre-configured in Microsoft Word with:
- RTL paragraph direction
- Right alignment
- Hebrew fonts
- Hebrew language settings (he-IL)
- Proper list styles for bullets

**Important**: All RTL formatting must be configured in the template. The Python code will NOT override these settings.

## Prerequisites

1. **Microsoft Word** (2016 or later recommended)
2. **Hebrew Editing Language** enabled in Word:
   - File → Options → Language
   - Under "Authoring Languages", add Hebrew (Israel)
   - Ensure Hebrew proofing tools are installed (optional but recommended)

## Step-by-Step Instructions

### Step 1: Create a New Document

1. Open Microsoft Word
2. Create a new blank document
3. Save it temporarily as a regular .docx file (we'll convert to template later)

### Step 2: Configure Normal Style

1. Open the **Styles** pane:
   - Home tab → Styles group → click the small arrow in bottom-right
   - Or press `Ctrl+Alt+Shift+S`

2. Right-click on **Normal** style → **Modify**

3. Configure the following settings:

   **Formatting:**
   - Font: Choose a Hebrew-supporting font (Arial, Times New Roman, or David)
   - Size: 12pt (or your preferred size)
   - Language: Hebrew (Israel) - `he-IL`
   
   **Paragraph Settings** (click Format → Paragraph):
   - Alignment: **Right**
   - Text direction: **Right-to-Left** (if available in your Word version)
   - Indentation: Set as needed (typically 0 for both left and right)
   - Spacing: Set line spacing and paragraph spacing as desired

4. Click **OK** to save Normal style

### Step 3: Configure List Bullet Style

1. In the Styles pane, find **List Bullet** style
   - If not visible, click "Options" in Styles pane and show all styles

2. Right-click **List Bullet** → **Modify**

3. Configure the following settings:

   **Formatting:**
   - Font: Same as Normal style (Hebrew-supporting font)
   - Size: Same as Normal style
   - Language: Hebrew (Israel) - `he-IL`
   
   **Paragraph Settings** (click Format → Paragraph):
   - Alignment: **Right**
   - Text direction: **Right-to-Left**
   - Bullet position: Ensure bullets appear on the right side
   - Indentation: Adjust so bullets align properly on the right

4. Click **OK** to save List Bullet style

### Step 4: Configure Heading Styles

Repeat for **Heading 1**, **Heading 2**, and **Heading 3**:

1. Right-click on **Heading 1** → **Modify**

2. Configure:
   - Font: Hebrew-supporting font (can be larger/bold for headings)
   - Size: Larger than Normal (e.g., 16pt for Heading 1)
   - Language: Hebrew (Israel) - `he-IL`
   - Paragraph → Alignment: **Right**
   - Paragraph → Text direction: **Right-to-Left**

3. Repeat for Heading 2 and Heading 3 (with appropriate sizes)

### Step 5: Set Document Defaults

1. Go to **File → Options → Language**

2. Under **Choose Editing Languages**:
   - Ensure Hebrew (Israel) is added
   - Set as default if desired

3. Go to **File → Options → Advanced**

4. Under **Show document content**:
   - Ensure "Right-to-left" options are enabled if available

### Step 6: Test the Template

Before saving as template, test it:

1. Type some Hebrew text: `זהו טקסט בעברית`
2. Verify:
   - Text aligns to the right
   - Text flows right-to-left
   - Font displays Hebrew characters correctly
   - Language is set to Hebrew (check via Review → Language)

3. Test a bullet list:
   - Type `- פריט ראשון` and press Enter
   - Verify bullet appears on the right side
   - Verify text is right-aligned

### Step 7: Save as Template

1. Go to **File → Save As**

2. Choose location (recommended: `templates/` folder in your project)

3. In **Save as type**, select **Word Template (*.dotx)**

4. Name it `hebrew_template.dotx`

5. Click **Save**

### Step 8: Verify Template

1. Close the template file

2. Create a new document from template:
   - File → New → Personal (or Custom)
   - Select your `hebrew_template.dotx`
   - Verify all styles work correctly

## Required Style Configurations Summary

| Style | Alignment | Text Direction | Language | Font |
|-------|-----------|----------------|----------|------|
| Normal | Right | RTL | he-IL | Hebrew-supporting |
| List Bullet | Right | RTL | he-IL | Hebrew-supporting |
| Heading 1 | Right | RTL | he-IL | Hebrew-supporting |
| Heading 2 | Right | RTL | he-IL | Hebrew-supporting |
| Heading 3 | Right | RTL | he-IL | Hebrew-supporting |

## Font Recommendations

**Recommended fonts** (widely available, good Hebrew support):
- **Arial**: Common, good Hebrew support
- **Times New Roman**: Classic, excellent Hebrew support
- **David**: Hebrew-specific font (if installed)

**Avoid**:
- Fonts that don't support Hebrew characters
- Fonts that may not be installed on target systems

## Troubleshooting

### Text Doesn't Align Right
- Check paragraph alignment is set to Right in style
- Verify text direction is Right-to-Left
- Ensure template is being loaded correctly

### Bullets Appear on Left
- Check List Bullet style alignment is Right
- Verify bullet position settings in paragraph formatting
- Test in Word directly before using in code

### Hebrew Characters Don't Display
- Ensure font supports Hebrew
- Check language is set to Hebrew (Israel) in style
- Verify Word has Hebrew editing language enabled

### Styles Not Applied
- Ensure template file is .dotx format (not .docx)
- Check that styles are properly saved in template
- Verify template path in code is correct

## Advanced Configuration (Optional)

### Custom Styles
You can create additional custom styles following the same pattern:
- Right alignment
- RTL text direction
- Hebrew language
- Hebrew-supporting font

### Headers and Footers
If you need headers/footers:
1. Insert → Header & Footer
2. Configure header/footer with RTL alignment
3. Set language to Hebrew
4. Save in template

### Page Setup
Configure page margins and layout:
- File → Page Setup
- Set margins as needed
- For RTL documents, you may want mirrored margins

## Verification Checklist

Before using the template with doc_builder, verify:

- [ ] Normal style: Right alignment, RTL direction, Hebrew language
- [ ] List Bullet style: Right alignment, RTL direction, bullets on right
- [ ] Heading styles: Right alignment, RTL direction, Hebrew language
- [ ] Document defaults: Hebrew language enabled
- [ ] Template saved as .dotx format
- [ ] Test document created from template works correctly
- [ ] Hebrew text displays correctly
- [ ] Bullets appear on right side

## Next Steps

Once template is created:

1. Place template in `templates/` folder
2. Update template path in your code:
   ```python
   create_doc_from_template(
       template_path='templates/hebrew_template.dotx',
       output_path='output.docx',
       text='טקסט בעברית'
   )
   ```
3. Test with simple Hebrew text
4. Verify output document has correct RTL formatting

## Support

If you encounter issues:
1. Verify all steps were followed correctly
2. Test template directly in Word before using in code
3. Check Word version compatibility
4. Ensure Hebrew editing language is enabled in Word
