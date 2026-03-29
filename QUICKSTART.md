# Quick Start Guide - 5 Minutes

Get started with the Bulk PowerPoint Generator in just 5 minutes!

## ⚡ Quick Setup

1. **Install Python** (if not already installed)
   - Download from python.org
   - ✅ Check "Add Python to PATH" during installation

2. **Install dependencies** - Open Command Prompt:
   ```cmd
   pip install pandas openpyxl
   ```

## 🎯 Create Your First Batch

### Step 1: Prepare Your Template (2 minutes)
1. Open PowerPoint
2. Create a slide with placeholders like this:
   ```
   Hello {{name}}!
   
   Your order #{{order_id}} for {{product}}
   has been confirmed.
   
   Total: ${{amount}}
   ```
3. Save as `template.pptx`

### Step 2: Create Your Data File (2 minutes)
Create a file named `data.csv`:
```csv
name,order_id,product,amount
John Smith,1001,Laptop,999
Jane Doe,1002,Monitor,299
Bob Johnson,1003,Keyboard,79
```

### Step 3: Run the Generator (1 minute)
1. Put both files in a folder (e.g., `C:\MyProject\`)
2. Put `bulk_pptx_generator.py` in the same folder
3. Open Command Prompt:
   ```cmd
   cd C:\MyProject
   python bulk_pptx_generator.py
   ```

4. Follow the prompts:
   - **Template**: Select `template.pptx`
   - **Data file**: Select `data.csv`
   - **Mapping**: 
     - `{{name}}` → 1 (or type "name")
     - `{{order_id}}` → 2 (or type "order_id")
     - `{{product}}` → 3 (or type "product")
     - `{{amount}}` → 4 (or type "amount")
   - **Filename format**: Type `{{name}} - Order {{order_id}}` (or press Enter for default)
   - **Output folder**: Press Enter (uses current folder) or specify a path
   - **Keep PPTX?**: Type `n` (just PDFs) or `y` (keep both)
   - **Confirm**: Type `y`

### Step 4: Get Your PDFs!
You'll find:
- `John Smith - Order 1001.pdf`
- `Jane Doe - Order 1002.pdf`
- `Bob Johnson - Order 1003.pdf`

## 🎓 Pro Tips

1. **Test First**: Always test with 1-2 rows before processing hundreds
2. **Column Names**: Use simple names in your CSV (no spaces or special characters work best)
3. **Custom Filenames**: Use placeholders in your filename format for organized output
   - `{{name}} Certificate` → "John Smith Certificate.pdf"
   - `{{date}} - {{client}} Invoice` → "2024-03-15 - ABC Corp Invoice.pdf"
   - `{{id}}_{{name}}_Report` → "1001_John Smith_Report.pdf"
4. **File Paths**: You can drag and drop files into Command Prompt to get their paths
5. **Errors**: If something goes wrong, check the error message - it usually tells you exactly what's wrong

## 🆘 Common First-Time Issues

**"Python is not recognized"**
→ You didn't check "Add Python to PATH" during installation. Reinstall Python.

**"No module named pandas"**
→ Run `pip install pandas openpyxl`

**"No placeholders found"**
→ Make sure you use `{{placeholder}}` with double curly braces

**"PowerPoint not found"**
→ This requires Microsoft PowerPoint to be installed

## 📱 Example Templates

### Certificate Template
```
╔════════════════════════════════╗
   CERTIFICATE OF ACHIEVEMENT
╚════════════════════════════════╝

Awarded to
{{student_name}}

For completing
{{course_name}}

Date: {{date}}
```

### Invoice Template
```
INVOICE

Bill To: {{client_name}}
Company: {{company}}
Date: {{invoice_date}}

Description: {{service}}
Amount: ${{amount}}

Thank you for your business!
```

### Name Badge Template
```
┌─────────────────────┐
│ HELLO                │
│ My name is           │
│                      │
│ {{name}}             │
│                      │
│ {{company}}          │
│ {{role}}             │
└─────────────────────┘
```

## ✅ Checklist

Before running the script, ensure:
- [ ] Python is installed
- [ ] `pandas` and `openpyxl` are installed (`pip install pandas openpyxl`)
- [ ] Microsoft PowerPoint is installed
- [ ] Your template has placeholders with `{{name}}`
- [ ] Your CSV/Excel file has column headers
- [ ] Template and data file are in an accessible folder

---

**Ready to automate?** Run `python bulk_pptx_generator.py` and follow the prompts!
