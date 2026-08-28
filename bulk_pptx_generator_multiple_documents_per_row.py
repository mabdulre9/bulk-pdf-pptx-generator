#!/usr/bin/env python3
"""
Bulk PowerPoint Generator - Windows Edition (Multi-Template)
Reads data from CSV/Excel and generates individual PDFs from one or more
PowerPoint templates. For each row of data, it generates a document from
EVERY configured template before moving on to the next row.

Requires: Microsoft PowerPoint installed on Windows
"""

import os
import sys
import re
from pathlib import Path
from dataclasses import dataclass, field
import pandas as pd
from typing import Dict, List, Optional
import shutil
import zipfile


# ============================================================
# Small helpers
# ============================================================

def clear_screen():
    """Clear the terminal screen."""
    os.system('cls')


def print_header(text: str):
    """Print a formatted header."""
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60 + "\n")


def get_file_path(prompt: str, file_types: List[str]) -> Path:
    """Get a valid file path from user."""
    while True:
        path_str = input(f"{prompt}\n> ").strip().strip('"').strip("'")
        if not path_str:
            print("❌ Path cannot be empty. Please try again.\n")
            continue

        path = Path(path_str).expanduser().resolve()

        if not path.exists():
            print(f"❌ File not found: {path}\n")
            continue

        if not path.is_file():
            print(f"❌ Path is not a file: {path}\n")
            continue

        if path.suffix.lower() not in file_types:
            print(f"❌ Invalid file type. Expected: {', '.join(file_types)}\n")
            continue

        return path


def get_directory_path(prompt: str, create_if_missing: bool = False) -> Path:
    """Get a valid directory path from user."""
    while True:
        path_str = input(f"{prompt}\n> ").strip().strip('"').strip("'")
        if not path_str:
            print("❌ Path cannot be empty. Please try again.\n")
            continue

        path = Path(path_str).expanduser().resolve()

        if not path.exists():
            if create_if_missing:
                try:
                    path.mkdir(parents=True, exist_ok=True)
                    print(f"✓ Created directory: {path}\n")
                    return path
                except Exception as e:
                    print(f"❌ Could not create directory: {e}\n")
                    continue
            else:
                print(f"❌ Directory not found: {path}\n")
                continue

        if not path.is_dir():
            print(f"❌ Path is not a directory: {path}\n")
            continue

        return path


def ask_yes_no(prompt: str, default: Optional[bool] = None) -> bool:
    """Ask a yes/no question. `default` (True/False/None) controls behavior on empty input."""
    suffix = " (y/n): "
    if default is True:
        suffix = " (Y/n): "
    elif default is False:
        suffix = " (y/N): "

    while True:
        ans = input(f"{prompt}{suffix}").strip().lower()
        if not ans and default is not None:
            return default
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("Please answer y or n.\n")


def load_data_file(file_path: Path) -> pd.DataFrame:
    """Load CSV or Excel file into a pandas DataFrame."""
    print_header(f"Loading data from: {file_path.name}")

    try:
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path)
        else:  # Excel
            df = pd.read_excel(file_path)

        print(f"✓ Successfully loaded {len(df)} rows\n")
        return df

    except Exception as e:
        print(f"❌ Error loading file: {e}")
        sys.exit(1)


# ============================================================
# Template config
# ============================================================

@dataclass
class TemplateConfig:
    """Everything needed to generate documents from one template."""
    name: str                      # short label, e.g. "Template 1"
    path: Path
    placeholders: List[str] = field(default_factory=list)
    mapping: Dict[str, str] = field(default_factory=dict)   # placeholder -> column
    filename_format: str = ""

    # per-run counters
    successful: int = 0
    pptx_only: int = 0
    failed: int = 0


# ============================================================
# Placeholder extraction / replacement
# ============================================================

def extract_placeholders(pptx_path: Path) -> List[str]:
    """Extract all unique placeholders from a PowerPoint template."""
    placeholders = set()
    placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')

    try:
        temp_dir = Path(os.environ.get('TEMP', 'C:\\Temp')) / f'pptx_temp_{pptx_path.stem}'

        if temp_dir.exists():
            shutil.rmtree(temp_dir)

        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        slides_dir = temp_dir / 'ppt' / 'slides'
        if slides_dir.exists():
            for slide_file in slides_dir.glob('slide*.xml'):
                content = slide_file.read_text(encoding='utf-8')
                # Strip XML tags so placeholders split across <a:t> runs are found
                text_only = re.sub(r'<[^>]+>', '', content)
                found = placeholder_pattern.findall(text_only)
                placeholders.update(ph.strip() for ph in found)

        shutil.rmtree(temp_dir)

        placeholder_list = sorted(placeholders)

        if placeholder_list:
            print("Found placeholders:")
            for ph in placeholder_list:
                print(f"  • {{{{{ph}}}}}")
            print()
        else:
            print("⚠️  No placeholders found in this template!")
            print("Make sure it contains placeholders like {{name}}, {{domain}}, etc.\n")

        return placeholder_list

    except Exception as e:
        print(f"❌ Error extracting placeholders: {e}")
        return []


def map_columns_to_placeholders(df: pd.DataFrame, placeholders: List[str]) -> Dict[str, str]:
    """Map DataFrame columns to template placeholders."""
    print("Available columns in your data file:")
    for i, col in enumerate(df.columns, 1):
        print(f"  {i}. {col}")
    print()

    mapping = {}

    for placeholder in placeholders:
        print(f"Which column should be used for {{{{{placeholder}}}}}?")

        while True:
            user_input = input(
                f"Enter column name or number (1-{len(df.columns)}), or press Enter to skip: "
            ).strip()

            if not user_input:
                print(f"⚠️  Skipping {{{{{placeholder}}}}} (will remain unchanged)\n")
                break

            try:
                col_num = int(user_input)
                if 1 <= col_num <= len(df.columns):
                    column_name = df.columns[col_num - 1]
                    mapping[placeholder] = column_name
                    print(f"✓ {{{{{placeholder}}}}} → {column_name}\n")
                    break
                else:
                    print(f"❌ Number must be between 1 and {len(df.columns)}\n")
            except ValueError:
                if user_input in df.columns:
                    mapping[placeholder] = user_input
                    print(f"✓ {{{{{placeholder}}}}} → {user_input}\n")
                    break
                else:
                    print(f"❌ Column '{user_input}' not found. Try again.\n")

    return mapping


def get_filename_format(mapping: Dict[str, str], template_label: str) -> str:
    """Get custom filename format from user for a specific template."""
    print(f"Specify how you want to name output files for {template_label}.")

    if mapping:
        print("Available placeholders:")
        for placeholder in mapping.keys():
            print(f"  • {{{{{placeholder}}}}}")
        print()

        placeholders_list = list(mapping.keys())
        print("Examples:")
        if len(placeholders_list) >= 2:
            print(f"  • {{{{{placeholders_list[0]}}}}} {{{{{placeholders_list[1]}}}}} Certificate")
            print(f"  • {{{{{placeholders_list[0]}}}}} - {{{{{placeholders_list[1]}}}}} Report")
        else:
            print(f"  • {{{{{placeholders_list[0]}}}}} Certificate")
        print()

    print("Enter your filename format (without .pdf extension):")
    print("Press Enter to use the first mapped column value as filename")

    filename_format = input("> ").strip()

    if filename_format:
        print(f"✓ Files from {template_label} will be named like: {filename_format}.pdf\n")
    else:
        if mapping:
            first_placeholder = list(mapping.keys())[0]
            print(f"✓ Files from {template_label} will be named using: {{{{{first_placeholder}}}}}.pdf\n")
        else:
            print(f"✓ Files from {template_label} will be named: document_1.pdf, document_2.pdf, etc.\n")

    return filename_format


def replace_placeholders_in_xml(xml_content: str, replacements: Dict[str, str]) -> str:
    """Replace placeholders in XML content, handling split placeholders and case sensitivity."""
    replacements_lower = {key.lower(): str(value) for key, value in replacements.items()}

    text_only = re.sub(r'<[^>]+>', '', xml_content)
    placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')
    found_placeholders = set(ph.strip() for ph in placeholder_pattern.findall(text_only))

    for placeholder in found_placeholders:
        replacement_value = replacements_lower.get(placeholder.lower())

        if replacement_value is not None:
            chars = list(placeholder)
            pattern_parts = []
            for char in chars:
                pattern_parts.append(re.escape(char) + r'(?:</[^>]+>(?:<[^>]+>)*)?')

            pattern = r'\{\{' + ''.join(pattern_parts) + r'\}\}'
            xml_content = re.sub(pattern, replacement_value, xml_content, flags=re.DOTALL)

    return xml_content


def generate_single_pptx(template_path: Path, output_path: Path,
                          replacements: Dict[str, str]) -> bool:
    """Generate a single PPTX file from template with replacements."""
    try:
        temp_base = Path(os.environ.get('TEMP', 'C:\\Temp'))
        temp_dir = temp_base / 'pptx_gen' / output_path.stem
        # add a bit of uniqueness so parallel/duplicate stems don't collide
        temp_dir = temp_dir.parent / f"{temp_dir.name}_{abs(hash(str(output_path)))}"

        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir(parents=True)

        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        slides_dir = temp_dir / 'ppt' / 'slides'
        if slides_dir.exists():
            for slide_file in slides_dir.glob('slide*.xml'):
                content = slide_file.read_text(encoding='utf-8')
                modified_content = replace_placeholders_in_xml(content, replacements)
                slide_file.write_text(modified_content, encoding='utf-8')

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zip_out.write(file_path, arcname)

        shutil.rmtree(temp_dir)
        return True

    except Exception as e:
        print(f"❌ Error generating PPTX: {e}")
        return False


def convert_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    """Convert PPTX to PDF using Microsoft PowerPoint via COM automation."""

    powerpoint = None
    deck = None

    try:
        import comtypes.client

        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

        # IMPORTANT: Do NOT set Visible = 0.
        # Some PowerPoint versions do not allow hiding the app window
        # during this COM operation.
        powerpoint.Visible = 1

        pptx_abs = str(pptx_path.absolute())
        pdf_abs = str(pdf_path.absolute())

        deck = powerpoint.Presentations.Open(pptx_abs, ReadOnly=True, WithWindow=True)

        # 32 = PDF format
        deck.SaveAs(pdf_abs, 32)

        deck.Close()
        deck = None

        powerpoint.Quit()
        powerpoint = None

        if pdf_path.exists():
            return True

        print("  ❌ PowerPoint reported success but PDF was not created.")
        return False

    except ImportError:
        print("  ❌ comtypes module not found. Run: pip install comtypes")
        return False

    except Exception as e:
        print(f"  ⚠️ PowerPoint conversion error: {e}")

        try:
            if deck is not None:
                deck.Close()
        except Exception:
            pass

        try:
            if powerpoint is not None:
                powerpoint.Quit()
        except Exception:
            pass

        return False


# ============================================================
# Setup: collect one or more templates
# ============================================================

def setup_templates(df: pd.DataFrame) -> List[TemplateConfig]:
    """Ask the user for one or more templates, extracting placeholders,
    mapping columns, and getting a filename format for each."""

    templates: List[TemplateConfig] = []
    template_num = 1

    while True:
        print_header(f"Template {template_num}: Select PowerPoint Template")
        template_path = get_file_path(
            "Enter the path to your PowerPoint template file (.pptx):",
            ['.pptx']
        )

        label = f"Template {template_num} ({template_path.name})"

        print_header(f"{label}: Extracting placeholders")
        placeholders = extract_placeholders(template_path)

        if not placeholders:
            print("⚠️  Warning: No placeholders found in this template.")
            if not ask_yes_no("Continue with this template anyway?", default=True):
                print("Skipping this template.\n")
                # Ask whether to try adding a template again, or stop entirely
                if ask_yes_no("Add a different template instead?", default=True):
                    continue
                else:
                    break

        print_header(f"{label}: Map Columns to Placeholders")
        mapping = {}
        if placeholders:
            mapping = map_columns_to_placeholders(df, placeholders)

        print_header(f"{label}: Filename Format")
        filename_format = get_filename_format(mapping, label)

        templates.append(TemplateConfig(
            name=label,
            path=template_path,
            placeholders=placeholders,
            mapping=mapping,
            filename_format=filename_format,
        ))

        print(f"✓ {label} configured.\n")

        if not ask_yes_no("Do you want to add another template?", default=False):
            break

        template_num += 1

    return templates


# ============================================================
# Filename building
# ============================================================

def build_filename(template: TemplateConfig, row: pd.Series, row_idx: int) -> str:
    """Build a sanitized output filename (no extension) for one row + template."""
    filename_base = template.filename_format

    if filename_base:
        for placeholder, column in template.mapping.items():
            value = row[column]
            if pd.isna(value):
                value = ""
            filename_base = filename_base.replace(f'{{{{{placeholder}}}}}', str(value))

    if not filename_base.strip() or '{{' in filename_base:
        if template.mapping:
            first_col = list(template.mapping.values())[0]
            filename_base = str(row[first_col])
        else:
            filename_base = f"document_{row_idx + 1}"

    # Clean filename - remove invalid characters
    filename_base = filename_base.replace('/', '_').replace('\\', '_')
    filename_base = re.sub(r'[<>:"|?*]', '_', filename_base)
    filename_base = filename_base.strip()
    filename_base = filename_base[:200]  # Limit length

    return filename_base


# ============================================================
# Main generation loop: ROW OUTER, TEMPLATE INNER
# ============================================================

def generate_bulk_documents(templates: List[TemplateConfig], df: pd.DataFrame,
                             output_dir: Path, save_pptx: bool = False):
    """For each row, generate a document from every template, then move to
    the next row."""

    total_docs = len(df) * len(templates)
    print_header(f"Generating {total_docs} documents "
                 f"({len(df)} rows x {len(templates)} template(s))")

    temp_pptx_dir = output_dir / 'temp_pptx'
    temp_pptx_dir.mkdir(exist_ok=True)

    # One subfolder per template if the user wants to keep PPTX files,
    # and/or if there's more than one template (avoids output collisions).
    pptx_dirs: Dict[str, Path] = {}
    out_dirs: Dict[str, Path] = {}
    multiple_templates = len(templates) > 1

    for template in templates:
        if save_pptx:
            pdir = (output_dir / 'pptx_files' if not multiple_templates
                    else output_dir / 'pptx_files' / _safe_dir_name(template.name))
            pdir.mkdir(parents=True, exist_ok=True)
            pptx_dirs[template.name] = pdir

        odir = output_dir if not multiple_templates else output_dir / _safe_dir_name(template.name)
        odir.mkdir(parents=True, exist_ok=True)
        out_dirs[template.name] = odir

    doc_counter = 0

    for row_idx, row in df.iterrows():
        print(f"\n--- Row {row_idx + 1}/{len(df)} ---")

        for template in templates:
            doc_counter += 1

            replacements = {}
            for placeholder, column in template.mapping.items():
                value = row[column]
                if pd.isna(value):
                    value = ""
                replacements[placeholder] = str(value)

            filename_base = build_filename(template, row, row_idx)

            target_out_dir = out_dirs[template.name]
            pptx_path = temp_pptx_dir / f"{_safe_dir_name(template.name)}__{filename_base}.pptx"
            pdf_path = target_out_dir / f"{filename_base}.pdf"

            print(f"[{doc_counter}/{total_docs}] {template.name}: {filename_base}")

            if generate_single_pptx(template.path, pptx_path, replacements):
                if convert_pptx_to_pdf(pptx_path, pdf_path):
                    print(f"  ✓ Created: {pdf_path.relative_to(output_dir)}")
                    template.successful += 1

                    if save_pptx:
                        shutil.copy2(pptx_path, pptx_dirs[template.name] / f"{filename_base}.pptx")
                else:
                    # PDF conversion failed, keep the PPTX somewhere useful
                    if save_pptx:
                        final_pptx = pptx_dirs[template.name] / f"{filename_base}.pptx"
                    else:
                        final_pptx = target_out_dir / f"{filename_base}.pptx"
                    shutil.copy2(pptx_path, final_pptx)
                    print(f"  ⚠️  PDF conversion failed, PPTX saved: {final_pptx.relative_to(output_dir)}")
                    template.pptx_only += 1
            else:
                print("  ❌ Failed to generate document")
                template.failed += 1

    print("\nCleaning up temporary files...")
    shutil.rmtree(temp_pptx_dir, ignore_errors=True)

    print_header("Generation Complete")
    for template in templates:
        print(f"{template.name}:")
        print(f"  ✓ Successful (PDF): {template.successful}")
        if template.pptx_only > 0:
            print(f"  ⚠️  PPTX only (PDF conversion failed): {template.pptx_only}")
        if template.failed > 0:
            print(f"  ❌ Failed: {template.failed}")
        print()

    print(f"Files saved under: {output_dir}")

    any_pptx_only = any(t.pptx_only > 0 for t in templates)
    if any_pptx_only:
        print("\n" + "!" * 60)
        print("MANUAL CONVERSION NEEDED:")
        print("Some files could not be converted to PDF automatically.")
        print("PPTX files have been saved. You can:")
        print("1. Open each PPTX in PowerPoint and 'Save As' PDF")
        print("2. Use an online converter like smallpdf.com")
        print("!" * 60)


def _safe_dir_name(name: str) -> str:
    """Turn a template label into a filesystem-safe folder name."""
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)
    return safe.strip()[:100]


# ============================================================
# PowerPoint availability check
# ============================================================

def check_powerpoint():
    """Check if PowerPoint is available and install comtypes if needed."""
    try:
        import comtypes.client
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Quit()
            return True
        except Exception:
            print("❌ Microsoft PowerPoint is not installed or not accessible.")
            return False
    except ImportError:
        print("⚠️  'comtypes' module not found. Installing...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "comtypes"])
            print("✓ comtypes installed successfully")
            return check_powerpoint()
        except Exception:
            print("❌ Could not install comtypes. Please install manually:")
            print("   pip install comtypes")
            return False


# ============================================================
# Main
# ============================================================

def main():
    clear_screen()
    print("╔════════════════════════════════════════════════════════════╗")
    print("║                                                            ║")
    print("║     BULK POWERPOINT TO PDF GENERATOR (MULTI-TEMPLATE)      ║")
    print("║     Windows Edition - Requires Microsoft PowerPoint        ║")
    print("║                                                            ║")
    print("╚════════════════════════════════════════════════════════════╝")

    print("\nChecking for Microsoft PowerPoint...")
    if not check_powerpoint():
        print("\n❌ This script requires Microsoft PowerPoint to be installed.")
        input("\nPress Enter to exit...")
        return

    print("✓ Microsoft PowerPoint detected\n")

    # Step 1: Data file first, since column mapping is needed for every template
    print_header("Step 1: Select Data File")
    data_path = get_file_path(
        "Enter the path to your CSV or Excel file:",
        ['.csv', '.xlsx', '.xls']
    )
    df = load_data_file(data_path)

    if df.empty:
        print("❌ Data file is empty!")
        return

    # Step 2: One or more templates
    print_header("Step 2: Configure Template(s)")
    templates = setup_templates(df)

    if not templates:
        print("❌ No templates configured. Exiting.")
        return

    # Step 3: Output directory
    print_header("Step 3: Select Output Directory")
    output_dir = get_directory_path(
        "Enter the directory where PDF files should be saved\n(will be created if it doesn't exist):",
        create_if_missing=True
    )

    # Step 4: Keep PPTX files?
    save_pptx = ask_yes_no("\nDo you want to keep the PPTX files as well?", default=False)

    # Step 5: Skip rows
    skip_input = input("\nHow many rows do you want to skip? (Press Enter for 0): ").strip()
    skip_rows = int(skip_input) if skip_input.isdigit() else 0

    if skip_rows > 0:
        df = df.iloc[skip_rows:].reset_index(drop=True)
        print(f"✓ Skipping {skip_rows} rows. Starting generation from row {skip_rows + 1}.")

    # Step 6: Confirm
    print_header("Ready to Generate")
    print(f"Data file: {data_path.name}")
    print(f"Rows to process: {len(df)}")
    print(f"Templates: {len(templates)}")
    for t in templates:
        print(f"\n  {t.name}")
        print(f"    File: {t.path}")
        if t.mapping:
            print("    Mappings:")
            for ph, col in t.mapping.items():
                print(f"      {{{{{ph}}}}} ← {col}")
        else:
            print("    Mappings: (none — template used as-is)")
        print(f"    Filename format: {t.filename_format or '(default)'}.pdf")

    print(f"\nOutput directory: {output_dir}")
    if len(templates) > 1:
        print("(Each template's PDFs will be saved in its own subfolder.)")
    print(f"\nTotal documents to generate: {len(df) * len(templates)}")
    print()

    if not ask_yes_no("Generate all documents?", default=True):
        print("Cancelled.")
        return

    # Step 7: Generate — row outer, template inner
    generate_bulk_documents(templates, df, output_dir, save_pptx)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)
