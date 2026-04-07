"""
restyle_pptx.py

Restyles AEM_Guides_Review_Feature.pptx so that it uses the same visual
styling and slide templates as LinkedIn_AEM_Guides_Review_Deck.pptx, while
keeping all the original content (text, diagrams, data) intact.

Approach:
- Replaces the slide master, all slide layouts, and theme from the LinkedIn
  deck into the content PPTX at the zip/XML level.
- Remaps every content slide's layout reference to the LinkedIn 'BLANK' layout
  (slideLayout2.xml), since all content slides use a blank/empty layout.
- Keeps all media files and slide XML from the content PPTX unchanged.
- Overwrites AEM_Guides_Review_Feature.pptx with the restyled version.
"""

import io
import shutil
import zipfile
import xml.etree.ElementTree as ET

CONTENT_FILE = "AEM_Guides_Review_Feature.pptx"
STYLE_FILE = "LinkedIn_AEM_Guides_Review_Deck.pptx"

# In the LinkedIn deck, slideLayout2 is "BLANK" — the best match for the
# content slides which all use a "Blank" layout.
LINKEDIN_BLANK_LAYOUT = "slideLayout2.xml"

# Files to copy verbatim from the LinkedIn deck into the restyled PPTX.
STYLE_MASTER_FILES = [
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/theme/theme1.xml",
]
for i in range(1, 12):
    STYLE_MASTER_FILES.append(f"ppt/slideLayouts/slideLayout{i}.xml")
    STYLE_MASTER_FILES.append(f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels")

SLIDE_LAYOUT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)


def build_slide_rels_xml(original_rels_xml: str, new_layout_target: str) -> str:
    """
    Return a slide .rels XML string where the slideLayout relationship is
    updated to point to *new_layout_target*, keeping all other relationships
    (images, notes slides, etc.) unchanged.
    """
    ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", ns)
    root = ET.fromstring(original_rels_xml)
    for rel in root:
        if rel.attrib.get("Type") == SLIDE_LAYOUT_REL_TYPE:
            rel.set("Target", f"../slideLayouts/{new_layout_target}")
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def restyle():
    # Read all files from the style (LinkedIn) deck into memory.
    with zipfile.ZipFile(STYLE_FILE, "r") as style_zip:
        style_files = {name: style_zip.read(name) for name in style_zip.namelist()}

    # Build the output PPTX in memory, based on the content file.
    output_buffer = io.BytesIO()

    with zipfile.ZipFile(CONTENT_FILE, "r") as content_zip, zipfile.ZipFile(
        output_buffer, "w", compression=zipfile.ZIP_DEFLATED
    ) as out_zip:

        for item in content_zip.infolist():
            name = item.filename

            if name in STYLE_MASTER_FILES:
                # Replace with the LinkedIn version
                out_zip.writestr(item, style_files[name])

            elif name.startswith("ppt/slides/_rels/") and name.endswith(".rels"):
                # Update the slideLayout reference in every slide's .rels file
                original = content_zip.read(name).decode("utf-8")
                updated = build_slide_rels_xml(original, LINKEDIN_BLANK_LAYOUT)
                # Prepend XML declaration to match original format
                updated_with_decl = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    + updated
                )
                out_zip.writestr(item, updated_with_decl.encode("utf-8"))

            else:
                # Copy everything else from the content file unchanged
                out_zip.writestr(item, content_zip.read(name))

    # Overwrite the content file with the restyled version
    with open(CONTENT_FILE, "wb") as f:
        f.write(output_buffer.getvalue())

    print(f"Restyled '{CONTENT_FILE}' successfully.")
    print(f"  Replaced {len(STYLE_MASTER_FILES)} master/layout/theme files from '{STYLE_FILE}'.")


if __name__ == "__main__":
    restyle()
