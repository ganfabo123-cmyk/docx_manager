
"""
docx_demo_generator.py

Demonstration script showing how to use docx_tools to build a
demo_extraction.json through a linear workflow.

Pipeline
--------
1. Load template extraction.json
2. Use docx_tools functions to append body_elements in order
3. Export demo_extraction.json
4. (Optional) compile into a .docx

Run
---
python docx_demo_generator.py
"""

import base64
from pathlib import Path

from docx_tools import (
    init_tools,
    HEADING1,
    HEADING2,
    TOC,
    SECTIONPTR,
    NORMAL,
    IMAGE,
    TABLE,
    OLE,
    OMATH,
    BUILD,
    COMPILE,
)

# ─────────────────────────────────────────────────────────────
# Utility
# ─────────────────────────────────────────────────────────────

def load_image_as_base64(path: str) -> str:
    """Read image file and convert to base64 string."""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


# ─────────────────────────────────────────────────────────────
# Demo workflow
# ─────────────────────────────────────────────────────────────

def generate_demo_document():

    # Initialize tools with template extraction.json
    init_tools("data/extraction.json")

    # ─────────────────────────────────────────
    # Cover Page
    # ─────────────────────────────────────────

    HEADING1("DOCX Engine V3: Structured Document Generation")

    NORMAL("Author: Demo User")
    NORMAL("Institution: Document Intelligence Lab")
    NORMAL("Date: 2026")

    SECTIONPTR("Abstract")

    # ─────────────────────────────────────────
    # Abstract
    # ─────────────────────────────────────────

    HEADING1("Abstract")

    NORMAL(
        "This paper presents DOCX Engine V3, a structured document generation "
        "framework designed to reconstruct complex Word documents from a "
        "linear JSON representation. The system introduces a novel pipeline "
        "that converts DOCX files into structured extraction.json data and "
        "rebuilds them through programmable tools."
    )

    NORMAL(
        "The proposed system separates document parsing, structural abstraction, "
        "and document reconstruction into independent stages, significantly "
        "improving maintainability and extensibility."
    )

    SECTIONPTR("Table of Contents")

    # ─────────────────────────────────────────
    # TOC
    # ─────────────────────────────────────────

    TOC()

    SECTIONPTR("Chapter 1")

    # ─────────────────────────────────────────
    # Chapter 1
    # ─────────────────────────────────────────

    HEADING1("Chapter 1 Introduction")

    NORMAL(
        "Document generation systems traditionally rely on template-based "
        "approaches. However, these systems struggle with complex formatting "
        "and dynamic content structures."
    )

    NORMAL(
        "DOCX Engine V3 addresses these issues by introducing a structured "
        "representation layer that separates content from formatting."
    )

    HEADING2("1.1 Motivation")

    NORMAL(
        "Modern document processing pipelines require high flexibility while "
        "preserving formatting fidelity. Existing tools often fail to achieve "
        "lossless reconstruction of DOCX files."
    )

    HEADING2("1.2 Contributions")

    NORMAL(
        "The main contributions of this work include a linear document "
        "representation model, a programmable tool abstraction layer, "
        "and a lossless DOCX reconstruction engine."
    )

    # ─────────────────────────────────────────
    # Image Demo
    # ─────────────────────────────────────────

    img_path = Path("data/demo_image.png")

    if img_path.exists():
        img_b64 = load_image_as_base64(img_path)
    else:
        img_b64 = ""

    IMAGE(
        img_b64,
        "Figure 1-1 Overall Architecture of DOCX Engine",
        position="center",
        width=140,
        height=90,
    )

    SECTIONPTR("Chapter 2")

    # ─────────────────────────────────────────
    # Chapter 2
    # ─────────────────────────────────────────

    HEADING1("Chapter 2 System Architecture")

    NORMAL(
        "The DOCX Engine V3 architecture consists of three primary modules: "
        "Extractor, Tool Layer, and Compiler."
    )

    HEADING2("2.1 Extractor")

    NORMAL(
        "The extractor parses DOCX XML structures and converts them into "
        "a structured JSON representation."
    )

    HEADING2("2.2 Tool Layer")

    NORMAL(
        "The tool layer provides programmable functions that construct "
        "document elements such as headings, tables, images, and formulas."
    )

    # ─────────────────────────────────────────
    # Table Demo
    # ─────────────────────────────────────────

    TABLE(
        [
            ["Module", "Description"],
            ["Extractor", "Parse DOCX into structured JSON"],
            ["Tool Layer", "Generate document elements"],
            ["Compiler", "Reconstruct DOCX from JSON"],
        ],
        "Table 2-1 System Modules",
    )

    SECTIONPTR("Chapter 3")

    # ─────────────────────────────────────────
    # Chapter 3
    # ─────────────────────────────────────────

    HEADING1("Chapter 3 Mathematical Representation")

    NORMAL(
        "Mathematical formulas are represented using the Office MathML "
        "standard supported by DOCX."
    )

    HEADING2("3.1 Example Formula")

    OMATH(
        "<m:oMath><m:r><m:t>E = mc^2</m:t></m:r></m:oMath>",
        "(3-1)",
    )

    NORMAL(
        "Equation (3-1) illustrates the mass-energy equivalence principle."
    )

    HEADING2("3.2 OLE Formula")

    OLE(
        base64_str="",
        formula_index="(3-2)"
    )

    NORMAL(
        "OLE formulas are used when equations are embedded through external "
        "applications such as Microsoft Equation Editor."
    )

    SECTIONPTR("References")

    # ─────────────────────────────────────────
    # References
    # ─────────────────────────────────────────

    HEADING1("References")

    NORMAL("[1] Leslie Lamport. LaTeX: A Document Preparation System.")

    NORMAL(
        "[2] ISO/IEC 29500. Office Open XML File Formats."
    )

    NORMAL(
        "[3] Knuth, Donald. The TeXbook."
    )

    # ─────────────────────────────────────────
    # Build extraction.json
    # ─────────────────────────────────────────

    BUILD("data/demo_extraction.json")

    print("\nDemo extraction.json generated.")

# ─────────────────────────────────────────────────────────────
# Entry
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":

    generate_demo_document()

    # Optional: compile to DOCX
    # COMPILE("demo_output.docx")

