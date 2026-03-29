#!/usr/bin/env python3
"""
Bulk PowerPoint Generator - Windows Edition
Reads data from CSV/Excel and generates individual PDFs from a PowerPoint template.
Requires: Microsoft PowerPoint installed on Windows
"""

import os
import sys
import re
from pathlib import Path
import pandas as pd
from typing import Dict, List
import shutil
import zipfile


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


def extract_placeholders(pptx_path: Path) -> List[str]:
    """Extract all unique placeholders from PowerPoint template."""
    print_header("Extracting placeholders from template")
    
    placeholders = set()
    placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')
    
    try:
        # Create temporary directory
        temp_dir = Path(os.environ.get('TEMP', 'C:\\Temp')) / 'pptx_temp'
        
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        
        # Unzip the PPTX file
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Search for placeholders in slide XML files
        slides_dir = temp_dir / 'ppt' / 'slides'
        if slides_dir.exists():
            for slide_file in slides_dir.glob('slide*.xml'):
                content = slide_file.read_text(encoding='utf-8')
                
                # Remove all XML tags to get continuous text
                # This handles split placeholders across multiple <a:t> tags
                text_only = re.sub(r'<[^>]+>', '', content)
                
                # Now find placeholders in the continuous text
                found = placeholder_pattern.findall(text_only)
                placeholders.update(ph.strip() for ph in found)
        
        # Clean up
        shutil.rmtree(temp_dir)
        
        placeholder_list = sorted(list(placeholders))
        
        if placeholder_list:
            print("Found placeholders:")
            for ph in placeholder_list:
                print(f"  • {{{{{ph}}}}}")
            print()
        else:
            print("⚠️  No placeholders found in template!")
            print("Make sure your template contains placeholders like {{name}}, {{domain}}, etc.\n")
        
        return placeholder_list
    
    except Exception as e:
        print(f"❌ Error extracting placeholders: {e}")
        return []


def map_columns_to_placeholders(df: pd.DataFrame, placeholders: List[str]) -> Dict[str, str]:
    """Map DataFrame columns to template placeholders."""
    print_header("Column to Placeholder Mapping")
    
    print("Available columns in your data file:")
    for i, col in enumerate(df.columns, 1):
        print(f"  {i}. {col}")
    print()
    
    mapping = {}
    
    for placeholder in placeholders:
        print(f"Which column should be used for {{{{{placeholder}}}}}?")
        
        while True:
            user_input = input(f"Enter column name or number (1-{len(df.columns)}), or press Enter to skip: ").strip()
            
            if not user_input:
                print(f"⚠️  Skipping {{{{{placeholder}}}}} (will remain unchanged)\n")
                break
            
            # Try to parse as number first
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
                # Try as column name
                if user_input in df.columns:
                    mapping[placeholder] = user_input
                    print(f"✓ {{{{{placeholder}}}}} → {user_input}\n")
                    break
                else:
                    print(f"❌ Column '{user_input}' not found. Try again.\n")
    
    return mapping


def get_filename_format(mapping: Dict[str, str]) -> str:
    """Get custom filename format from user."""
    print_header("Filename Format Configuration")
    
    print("Specify how you want to name your output files.")
    print("You can use placeholders and add custom text.\n")
    
    if mapping:
        print("Available placeholders:")
        for placeholder in mapping.keys():
            print(f"  • {{{{{placeholder}}}}}")
        print()
        
        print("Examples:")
        placeholders_list = list(mapping.keys())
        if len(placeholders_list) >= 2:
            print(f"  • {{{{{placeholders_list[0]}}}}} {{{{{placeholders_list[1]}}}}} Certificate")
            print(f"  • {{{{{placeholders_list[0]}}}}} - {{{{{placeholders_list[1]}}}}} Report")
            print(f"  • Invoice_{{{{{placeholders_list[0]}}}}}")
        else:
            print(f"  • {{{{{placeholders_list[0]}}}}} Certificate")
            print(f"  • Report - {{{{{placeholders_list[0]}}}}}")
        print()
    
    print("Enter your filename format (without .pdf extension):")
    print("Press Enter to use the first column value as filename")
    
    filename_format = input("> ").strip()
    
    if filename_format:
        print(f"\n✓ Files will be named: {filename_format}.pdf")
        print(f"  Example: ", end="")
        # Show an example with placeholder names
        example = filename_format
        for placeholder in mapping.keys():
            example = example.replace(f'{{{{{placeholder}}}}}', placeholder)
        print(f"{example}.pdf\n")
    else:
        if mapping:
            first_placeholder = list(mapping.keys())[0]
            print(f"\n✓ Files will be named using: {{{{{first_placeholder}}}}}.pdf\n")
        else:
            print("\n✓ Files will be named: document_1.pdf, document_2.pdf, etc.\n")
    
    return filename_format

def replace_placeholders_in_xml(xml_content: str, replacements: Dict[str, str]) -> str:
    """Replace placeholders in XML content, handling split placeholders and case sensitivity."""
    
    # Create a case-insensitive version of replacements
    # Map lowercase placeholder names to their values
    replacements_lower = {}
    for key, value in replacements.items():
        replacements_lower[key.lower()] = str(value)
    
    # Find all placeholders in the XML (even if split across tags)
    # First, remove XML tags to get continuous text
    text_only = re.sub(r'<[^>]+>', '', xml_content)
    
    # Find all {{placeholder}} patterns
    placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')
    found_placeholders = set(ph.strip() for ph in placeholder_pattern.findall(text_only))
    
    # For each found placeholder, replace it in the XML
    for placeholder in found_placeholders:
        # Check if we have a replacement for this placeholder (case-insensitive)
        replacement_value = replacements_lower.get(placeholder.lower())
        
        if replacement_value is not None:
            # Replace all occurrences of {{placeholder}} with the value
            # Use a regex that handles the placeholder potentially split across tags
            
            # Build a flexible pattern that matches the placeholder even if split
            chars = list(placeholder)
            pattern_parts = []
            for char in chars:
                # Each character might be followed by XML tags
                pattern_parts.append(re.escape(char) + r'(?:</[^>]+>(?:<[^>]+>)*)?')
            
            pattern = r'\{\{' + ''.join(pattern_parts) + r'\}\}'
            
            xml_content = re.sub(pattern, replacement_value, xml_content, flags=re.DOTALL)
    
    return xml_content


def generate_single_pptx(template_path: Path, output_path: Path, 
                        replacements: Dict[str, str]) -> bool:
    """Generate a single PPTX file from template with replacements."""
    
    try:
        # Create temporary directory
        temp_base = Path(os.environ.get('TEMP', 'C:\\Temp'))
        temp_dir = temp_base / 'pptx_gen' / output_path.stem
        
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir(parents=True)
        
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Process all slide XML files
        slides_dir = temp_dir / 'ppt' / 'slides'
        if slides_dir.exists():
            for slide_file in slides_dir.glob('slide*.xml'):
                content = slide_file.read_text(encoding='utf-8')
                modified_content = replace_placeholders_in_xml(content, replacements)
                slide_file.write_text(modified_content, encoding='utf-8')
        
        # Repack as PPTX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zip_out.write(file_path, arcname)
        
        # Clean up
        shutil.rmtree(temp_dir)
        
        return True
    
    except Exception as e:
        print(f"❌ Error generating PPTX: {e}")
        return False


def convert_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    """Convert PPTX to PDF using Microsoft PowerPoint."""
    try:
        import comtypes.client
        
        # Create PowerPoint application
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        
        # Convert paths to absolute paths
        pptx_abs = str(pptx_path.absolute())
        pdf_abs = str(pdf_path.absolute())
        
        # Open presentation
        deck = powerpoint.Presentations.Open(pptx_abs)
        
        # Save as PDF (format code 32 = PDF)
        deck.SaveAs(pdf_abs, 32)
        deck.Close()
        powerpoint.Quit()
        
        return pdf_path.exists()
    
    except ImportError:
        print("  ❌ comtypes module not found. Installing...")
        return False
    except Exception as e:
        print(f"  ⚠️  PowerPoint conversion error: {e}")
        return False


def generate_bulk_documents(template_path: Path, df: pd.DataFrame, 
                           mapping: Dict[str, str], output_dir: Path,
                           filename_format: str, save_pptx: bool = False):
    """Generate all documents from the data."""
    
    print_header(f"Generating {len(df)} documents")
    
    # Create temporary PPTX directory
    temp_pptx_dir = output_dir / 'temp_pptx'
    temp_pptx_dir.mkdir(exist_ok=True)
    
    # Optionally create permanent PPTX directory
    pptx_dir = None
    if save_pptx:
        pptx_dir = output_dir / 'pptx_files'
        pptx_dir.mkdir(exist_ok=True)
    
    successful = 0
    failed = 0
    pptx_only = 0
    
    for idx, row in df.iterrows():
        # Create replacements dictionary for this row
        replacements = {}
        for placeholder, column in mapping.items():
            value = row[column]
            # Handle NaN values
            if pd.isna(value):
                value = ""
            replacements[placeholder] = str(value)
        
        # Generate filename based on user's format
        filename_base = filename_format
        
        # Replace placeholders in filename
        for placeholder, column in mapping.items():
            value = row[column]
            if pd.isna(value):
                value = ""
            # Replace placeholder in filename
            filename_base = filename_base.replace(f'{{{{{placeholder}}}}}', str(value))
        
        # If no format specified or still has unreplaced placeholders, use fallback
        if not filename_base.strip() or '{{' in filename_base:
            if mapping:
                first_col = list(mapping.values())[0]
                filename_base = str(row[first_col])
            else:
                filename_base = f"document_{idx + 1}"
        
        # Clean filename - remove invalid characters
        filename_base = filename_base.replace('/', '_').replace('\\', '_')
        filename_base = re.sub(r'[<>:"|?*]', '_', filename_base)
        filename_base = filename_base.strip()
        filename_base = filename_base[:200]  # Limit length
        
        pptx_path = temp_pptx_dir / f"{filename_base}.pptx"
        pdf_path = output_dir / f"{filename_base}.pdf"
        
        print(f"[{idx + 1}/{len(df)}] Generating: {filename_base}")
        
        # Generate PPTX
        if generate_single_pptx(template_path, pptx_path, replacements):
            # Try to convert to PDF
            if convert_pptx_to_pdf(pptx_path, pdf_path):
                print(f"  ✓ Created: {pdf_path.name}")
                successful += 1
                
                # If user wants to keep PPTX files, copy them
                if save_pptx and pptx_dir:
                    shutil.copy2(pptx_path, pptx_dir / f"{filename_base}.pptx")
            else:
                # PDF conversion failed, but PPTX was created
                if save_pptx and pptx_dir:
                    final_pptx = pptx_dir / f"{filename_base}.pptx"
                    shutil.copy2(pptx_path, final_pptx)
                    print(f"  ⚠️  PDF conversion failed, PPTX saved: {final_pptx.name}")
                else:
                    # Keep PPTX in temp for manual conversion
                    final_pptx = output_dir / f"{filename_base}.pptx"
                    shutil.copy2(pptx_path, final_pptx)
                    print(f"  ⚠️  PDF conversion failed, PPTX saved: {final_pptx.name}")
                pptx_only += 1
        else:
            print(f"  ❌ Failed to generate document")
            failed += 1
    
    # Clean up temporary PPTX files
    print("\nCleaning up temporary files...")
    shutil.rmtree(temp_pptx_dir)
    
    print_header("Generation Complete")
    print(f"✓ Successful (PDF): {successful}")
    if pptx_only > 0:
        print(f"⚠️  PPTX only (PDF conversion failed): {pptx_only}")
    print(f"❌ Failed: {failed}")
    print(f"\nFiles saved to: {output_dir}")
    
    if pptx_only > 0:
        print("\n" + "!" * 60)
        print("MANUAL CONVERSION NEEDED:")
        print("Some files could not be converted to PDF automatically.")
        print("PPTX files have been saved. You can:")
        print("1. Open each PPTX in PowerPoint and 'Save As' PDF")
        print("2. Use an online converter like smallpdf.com")
        print("!" * 60)


def check_powerpoint():
    """Check if PowerPoint is available and install comtypes if needed."""
    try:
        import comtypes.client
        # Try to create PowerPoint object
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Quit()
            return True
        except:
            print("❌ Microsoft PowerPoint is not installed or not accessible.")
            return False
    except ImportError:
        print("⚠️  'comtypes' module not found. Installing...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "comtypes"])
            print("✓ comtypes installed successfully")
            return check_powerpoint()  # Retry after installation
        except:
            print("❌ Could not install comtypes. Please install manually:")
            print("   pip install comtypes")
            return False


def main():
    """Main program flow."""
    
    clear_screen()
    print("╔════════════════════════════════════════════════════════════╗")
    print("║                                                            ║")
    print("║     BULK POWERPOINT TO PDF GENERATOR                       ║")
    print("║     Windows Edition - Requires Microsoft PowerPoint       ║")
    print("║                                                            ║")
    print("╚════════════════════════════════════════════════════════════╝")
    
    # Check for PowerPoint
    print("\nChecking for Microsoft PowerPoint...")
    if not check_powerpoint():
        print("\n❌ This script requires Microsoft PowerPoint to be installed.")
        print("Please install PowerPoint and try again.")
        input("\nPress Enter to exit...")
        return
    
    print("✓ Microsoft PowerPoint detected\n")
    
    # Step 1: Get template file
    print_header("Step 1: Select PowerPoint Template")
    template_path = get_file_path(
        "Enter the path to your PowerPoint template file (.pptx):",
        ['.pptx']
    )
    
    # Step 2: Extract placeholders
    placeholders = extract_placeholders(template_path)
    
    if not placeholders:
        print("⚠️  Warning: No placeholders found in template.")
        print("Continuing anyway - you can still generate documents.\n")
        cont = input("Continue? (y/n): ").strip().lower()
        if cont != 'y':
            print("Exiting...")
            return
    
    # Step 3: Get data file
    print_header("Step 2: Select Data File")
    data_path = get_file_path(
        "Enter the path to your CSV or Excel file:",
        ['.csv', '.xlsx', '.xls']
    )
    
    # Step 4: Load data
    df = load_data_file(data_path)
    
    if df.empty:
        print("❌ Data file is empty!")
        return
    
    # Step 5: Map columns to placeholders
    mapping = {}
    if placeholders:
        mapping = map_columns_to_placeholders(df, placeholders)
    
    # Step 6: Get filename format
    filename_format = get_filename_format(mapping)
    
    # Step 7: Get output directory
    print_header("Step 3: Select Output Directory")
    output_dir = get_directory_path(
        "Enter the directory where PDF files should be saved\n(will be created if it doesn't exist):",
        create_if_missing=True
    )
    
    # Step 8: Ask about saving PPTX files
    save_pptx_input = input("\nDo you want to keep the PPTX files as well? (y/n): ").strip().lower()
    save_pptx = save_pptx_input == 'y'
    
    # Step 9: Confirm and generate
    print_header("Ready to Generate")
    print(f"Template: {template_path.name}")
    print(f"Data file: {data_path.name}")
    print(f"Records: {len(df)}")
    print(f"Output: {output_dir}")
    print(f"\nMappings:")
    if mapping:
        for ph, col in mapping.items():
            print(f"  {{{{{ph}}}}} ← {col}")
    else:
        print("  (No mappings - template will be used as-is)")
    
    if filename_format:
        print(f"\nFilename format: {filename_format}.pdf")
    else:
        if mapping:
            first_col = list(mapping.values())[0]
            print(f"\nFilename format: {first_col}.pdf")
        else:
            print(f"\nFilename format: document_N.pdf")
    print()
    
    confirm = input("Generate all documents? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Cancelled.")
        return
    
    # Step 10: Generate documents
    generate_bulk_documents(template_path, df, mapping, output_dir, filename_format, save_pptx)
    
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
