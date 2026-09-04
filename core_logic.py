#!/usr/bin/env python3
"""
Core logic for Bulk PowerPoint Generator.
Handles template processing, placeholder replacement, and PDF conversion.
"""

import os
import re
import shutil
import zipfile
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import pandas as pd

@dataclass
class TemplateConfig:
    """Everything needed to generate documents from one template."""
    name: str                      # short label, e.g. "Template 1"
    path: Path
    placeholders: List[str] = field(default_factory=list)
    mapping: Dict[str, str] = field(default_factory=dict)   # placeholder -> column
    filename_format: str = ""
    custom_output_dir: Optional[Path] = None

    # per-run counters
    successful: int = 0
    pptx_only: int = 0
    failed: int = 0

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
                text_only = re.sub(r'<[^>]+>', '', content)
                found = placeholder_pattern.findall(text_only)
                placeholders.update(ph.strip() for ph in found)

        shutil.rmtree(temp_dir)
        return sorted(placeholders)

    except Exception:
        return []

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

def generate_single_pptx(template_path: Path, output_path: Path, replacements: Dict[str, str]) -> bool:
    """Generate a single PPTX file from template with replacements."""
    try:
        temp_base = Path(os.environ.get('TEMP', 'C:\\Temp'))
        temp_dir = temp_base / 'pptx_gen' / output_path.stem
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

    except Exception:
        return False

def convert_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    """Convert PPTX to PDF using Microsoft PowerPoint via COM automation."""
    powerpoint = None
    deck = None
    try:
        import comtypes.client
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        pptx_abs = str(pptx_path.absolute())
        pdf_abs = str(pdf_path.absolute())
        deck = powerpoint.Presentations.Open(pptx_abs, ReadOnly=True, WithWindow=True)
        deck.SaveAs(pdf_abs, 32)
        deck.Close()
        deck = None
        powerpoint.Quit()
        powerpoint = None
        return pdf_path.exists()
    except Exception:
        if deck: deck.Close()
        if powerpoint: powerpoint.Quit()
        return False

def build_filename(template: TemplateConfig, row: pd.Series, row_idx: int) -> str:
    """Build a sanitized output filename (no extension) for one row + template."""
    filename_base = template.filename_format
    if filename_base:
        for placeholder, column in template.mapping.items():
            value = row[column]
            if pd.isna(value): value = ""
            filename_base = filename_base.replace(f'{{{{{placeholder}}}}}', str(value))

    if not filename_base.strip() or '{{' in filename_base:
        if template.mapping:
            first_col = list(template.mapping.values())[0]
            filename_base = str(row[first_col])
        else:
            filename_base = f"document_{row_idx + 1}"

    filename_base = filename_base.replace('/', '_').replace('\\', '_')
    filename_base = re.sub(r'[<>:"|?*]', '_', filename_base)
    filename_base = filename_base.strip()[:200]
    return filename_base

def _safe_dir_name(name: str) -> str:
    """Turn a template label into a filesystem-safe folder name."""
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)
    return safe.strip()[:100]
def generate_bulk_documents(templates: List[TemplateConfig], df: pd.DataFrame, output_dir: Path, save_pptx: bool = False, separate_folders: bool = True, progress_callback=None):
    """For each row, generate a document from every template, then move to the next row."""
    total_docs = len(df) * len(templates)
    temp_pptx_dir = output_dir / 'temp_pptx'
    temp_pptx_dir.mkdir(exist_ok=True)

    pptx_dirs: Dict[str, Path] = {}
    out_dirs: Dict[str, Path] = {}

    for template in templates:
        if template.custom_output_dir:
            odir = template.custom_output_dir
        elif separate_folders:
            odir = output_dir / _safe_dir_name(template.name)
        else:
            odir = output_dir
        
        odir.mkdir(parents=True, exist_ok=True)
        out_dirs[template.name] = odir

        if save_pptx:
            pdir = odir / 'pptx_files'
            pdir.mkdir(parents=True, exist_ok=True)
            pptx_dirs[template.name] = pdir

    doc_counter = 0
    for row_idx, row in df.iterrows():
        for template in templates:
            doc_counter += 1
            replacements = {}
            for placeholder, column in template.mapping.items():
                value = row[column]
                if pd.isna(value): value = ""
                replacements[placeholder] = str(value)

            filename_base = build_filename(template, row, row_idx)
            target_out_dir = out_dirs[template.name]
            pptx_path = temp_pptx_dir / f"{_safe_dir_name(template.name)}__{filename_base}.pptx"
            pdf_path = target_out_dir / f"{filename_base}.pdf"

            if progress_callback:
                # More detailed message for the progress bar
                msg = f"Row {row_idx + 1}/{len(df)} | Template: {template.name} | File: {filename_base}.pdf"
                progress_callback(doc_counter, total_docs, msg)

            if generate_single_pptx(template.path, pptx_path, replacements):
                if convert_pptx_to_pdf(pptx_path, pdf_path):
                    template.successful += 1
                    if save_pptx:
                        shutil.copy2(pptx_path, pptx_dirs[template.name] / f"{filename_base}.pptx")
                else:
                    if save_pptx:
                        final_pptx = pptx_dirs[template.name] / f"{filename_base}.pptx"
                    else:
                        final_pptx = target_out_dir / f"{filename_base}.pptx"
                    shutil.copy2(pptx_path, final_pptx)
                    template.pptx_only += 1
            else:
                template.failed += 1

    shutil.rmtree(temp_pptx_dir, ignore_errors=True)
    return templates
