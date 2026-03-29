# Bulk PowerPoint to PDF Generator - Windows Edition

A Python tool that automates the generation of personalized PowerPoint presentations and PDFs from a template using data from CSV or Excel files.

## Overview

This tool allows you to:
- Create multiple personalized PowerPoint presentations from a single template
- Automatically replace placeholders (like `{{name}}`, `{{company}}`) with data from your spreadsheet
- Convert presentations to PDF automatically using Microsoft PowerPoint
- Process hundreds of documents in minutes instead of hours

## Use Cases

- **Certificates**: Generate personalized certificates for course participants
- **Reports**: Create individual reports for clients or departments
- **Invitations**: Mass-produce customized event invitations
- **Business Cards**: Generate multiple business cards with different details
- **Presentations**: Create personalized pitch decks for different prospects

## Requirements

### System Requirements
- **Operating System**: Windows 10 or Windows 11
- **Microsoft PowerPoint**: Must be installed (part of Microsoft Office)
- **Python**: Version 3.9 or higher

### Python Dependencies
- `pandas` - For reading CSV/Excel files
- `openpyxl` - For Excel file support
- `comtypes` - For PowerPoint automation (auto-installed by the script)

## Installation

### Step 1: Install Python
1. Download Python from [python.org](https://www.python.org/downloads/)
2. During installation, **check "Add Python to PATH"**
3. Verify installation by opening Command Prompt and typing:
   ```cmd
   python --version
   ```

### Step 2: Install Required Packages
Open Command Prompt and run:
```cmd
pip install pandas openpyxl
```

The `comtypes` package will be automatically installed when you first run the script if not already present.

### Step 3: Download the Script
Save `bulk_pptx_generator.py` to a folder on your computer (e.g., `C:\Scripts\`)

## How to Use

### Prepare Your Files

#### 1. Create a PowerPoint Template
Create a PowerPoint presentation with placeholders using double curly braces:
- `{{name}}` - Will be replaced with data from your spreadsheet
- `{{company}}` - Will be replaced with company names
- `{{date}}` - Will be replaced with dates
- `{{amount}}` - Will be replaced with numbers
- etc.

**Example slide:**
```
Certificate of Achievement

This is to certify that {{name}}
has successfully completed {{course}}
on {{date}}.
```

**Important Tips:**
- Placeholders must use double curly braces: `{{placeholder}}`
- Placeholders can only contain letters, numbers, and underscores
- Placeholders are case-sensitive: `{{Name}}` is different from `{{name}}`

#### 2. Create a Data File (CSV or Excel)

**CSV Example (data.csv):**
```csv
name,company,date,amount
John Smith,ABC Corp,2024-03-15,1500
Jane Doe,XYZ Inc,2024-03-16,2000
Bob Johnson,Tech Co,2024-03-17,1750
```

**Excel Example:**
| name         | company   | date       | amount |
|--------------|-----------|------------|--------|
| John Smith   | ABC Corp  | 2024-03-15 | 1500   |
| Jane Doe     | XYZ Inc   | 2024-03-16 | 2000   |
| Bob Johnson  | Tech Co   | 2024-03-17 | 1750   |

### Run the Script

1. **Open Command Prompt**
   - Press `Win + R`
   - Type `cmd` and press Enter

2. **Navigate to the script folder**
   ```cmd
   cd C:\Scripts
   ```

3. **Run the script**
   ```cmd
   python bulk_pptx_generator.py
   ```

4. **Follow the prompts:**
   - Select your PowerPoint template
   - Select your data file (CSV or Excel)
   - Map columns to placeholders
   - Choose output directory
   - Confirm and generate

### Interactive Mapping Process

The script will guide you through mapping your spreadsheet columns to template placeholders:

```
Which column should be used for {{name}}?
Enter column name or number (1-5), or press Enter to skip: 1
✓ {{name}} → name

Which column should be used for {{company}}?
Enter column name or number (1-5), or press Enter to skip: company
✓ {{company}} → company
```

You can:
- Enter the column number (e.g., `1`, `2`, `3`)
- Enter the column name (e.g., `name`, `company`)
- Press Enter to skip a placeholder

### Custom Filename Format

After mapping, you'll be asked how to name your output files:

```
Specify how you want to name your output files.
You can use placeholders and add custom text.

Available placeholders:
  • {{name}}
  • {{company}}
  • {{date}}

Examples:
  • {{name}} {{company}} Certificate
  • {{name}} - {{date}} Report
  • Invoice_{{company}}

Enter your filename format (without .pdf extension):
> {{name}} {{company}} Internship Certificate
```

**Filename Examples:**
- Input: `{{name}} Certificate` → Output: `John Smith Certificate.pdf`
- Input: `{{company}} - {{name}} Report` → Output: `ABC Corp - John Smith Report.pdf`
- Input: `Invoice_{{date}}_{{name}}` → Output: `Invoice_2024-03-15_John Smith.pdf`
- Press Enter to use the first column as filename

## 📁 Output

The script will create:
- **PDF files**: One for each row in your data file (in the output directory)
- **PPTX files** (optional): If you choose to keep them, stored in `pptx_files` subfolder

Files are named using the first mapped column value (e.g., `John Smith.pdf`, `Jane Doe.pdf`)

## 🔧 Troubleshooting

### "Microsoft PowerPoint is not installed or not accessible"
**Solution**: 
- Ensure Microsoft PowerPoint is installed
- Try running Command Prompt as Administrator
- Check if PowerPoint opens normally

### "comtypes module not found"
**Solution**: The script will try to install it automatically. If it fails:
```cmd
pip install comtypes
```

### "Error loading file"
**Solutions**:
- Ensure your CSV/Excel file is not open in another program
- Check that the file is not corrupted
- Verify the file path is correct

### "No placeholders found in template"
**Solutions**:
- Check that placeholders use double curly braces: `{{name}}`
- Ensure placeholders are in text boxes (not in notes or hidden areas)
- Verify the template is a `.pptx` file, not `.ppt`

### PDF conversion fails
**Solutions**:
- Close any open PowerPoint windows
- Ensure you have write permissions to the output directory
- Try running the script as Administrator
- Check that PowerPoint is properly activated

### Slow performance
**Tips**:
- PowerPoint opens and closes for each conversion - this is normal
- For 100 documents, expect 5-15 minutes depending on your computer
- Close other applications to free up memory
- Don't use your computer for other tasks while the script runs

## 💡 Tips and Best Practices

### Template Design
1. **Test with one record first**: Create a test data file with one row to verify your template
2. **Use simple placeholders**: Short, descriptive names like `{{name}}` instead of `{{participant_full_name}}`
3. **Keep formatting simple**: Complex animations or transitions may slow conversion
4. **Text boxes work best**: Placeholders in text boxes are most reliable

### Data File Preparation
1. **Clean your data**: Remove extra spaces, special characters
2. **Use consistent formats**: Dates, numbers should be formatted consistently
3. **Handle missing data**: The script will replace missing values with empty strings
4. **Check column names**: Column names should match what you expect to map

### Filename Best Practices
1. **Keep it simple**: Short, descriptive filenames work best
2. **Use unique identifiers**: Include at least one unique field (like name or ID)
3. **Avoid special characters**: The script will replace invalid characters with underscores
4. **Test first**: Verify your filename format with a small batch first
5. **Examples of good formats**:
   - `{{name}} - {{course}} Certificate`
   - `Invoice_{{invoice_id}}_{{client}}`
   - `{{last_name}}_{{first_name}}_Report`
   - `{{date}} - {{company}} - {{type}}`

### Performance Optimization
1. **Batch processing**: For very large datasets (1000+ records), consider splitting into smaller batches
2. **Simple templates**: Templates with fewer slides and simpler formatting convert faster
3. **Close other programs**: Free up system resources for PowerPoint

## 📝 Example Workflow

### Certificate Generation Example

**1. Template (certificate_template.pptx):**
```
╔═══════════════════════════════════════╗
║    CERTIFICATE OF COMPLETION          ║
╚═══════════════════════════════════════╝

This certifies that

{{participant_name}}

has successfully completed

{{course_name}}

on {{completion_date}}

Course ID: {{course_id}}
```

**2. Data file (participants.csv):**
```csv
participant_name,course_name,completion_date,course_id
Alice Johnson,Python Programming,March 15 2024,PY-101
Bob Smith,Data Analysis,March 16 2024,DA-202
Carol White,Web Development,March 17 2024,WD-303
```

**3. Run the script:**
```cmd
python bulk_pptx_generator.py
```

**4. Map the columns:**
- `{{participant_name}}` → participant_name
- `{{course_name}}` → course_name
- `{{completion_date}}` → completion_date
- `{{course_id}}` → course_id

**5. Set filename format:**
```
Enter your filename format: {{participant_name}} - {{course_name}} Certificate
```

**6. Output:**
- `Alice Johnson - Python Programming Certificate.pdf`
- `Bob Smith - Data Analysis Certificate.pdf`
- `Carol White - Web Development Certificate.pdf`

## ❓ FAQ

**Q: Can I use this on Mac or Linux?**
A: No, this version is specifically for Windows with Microsoft PowerPoint. For Mac/Linux, you would need a different version using LibreOffice.

**Q: How do I customize the output filenames?**
A: The script will ask you for a filename format after column mapping. You can use any combination of placeholders and text. For example: `{{name}} {{domain}} Internship Certificate` will create files like "John Smith Marketing Internship Certificate.pdf".

**Q: What if two people have the same name?**
A: Use multiple placeholders in your filename format to make them unique, like `{{name}} - {{id}} Certificate` or `{{name}} {{company}} Report`.

**Q: Can I include images in my template?**
A: Yes! Images in your template will be preserved. However, dynamic image replacement is not supported.

**Q: What happens if a column has empty values?**
A: Empty values will be replaced with blank text in the generated documents.

**Q: Can I use special characters in placeholders?**
A: Placeholders can only contain letters, numbers, and underscores. `{{name_1}}` works, but `{{name-1}}` or `{{name.1}}` don't.

**Q: How many documents can I generate at once?**
A: There's no hard limit, but for very large batches (1000+), consider splitting into smaller batches for better performance.

**Q: Can I customize the output filename?**
A: Currently, the filename is based on the first mapped column. You can modify the script if you need custom naming.

## 📞 Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Ensure all requirements are met
3. Verify your template and data file are correctly formatted
4. Try with a simple test case (1-2 records) first

## License

This script is provided as-is for educational and commercial use.

**Made with ❤️ for automating repetitive tasks**
