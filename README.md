# 📄 Bulk PowerPoint & PDF Generator

An automated tool for generating personalized PDF documents from PowerPoint templates using data from CSV or Excel files. This tool allows for mass-production of certificates, reports, or personalized presentations by mapping data columns to template placeholders.

## ✨ Features

- **Multi-Template Support**: Generate documents from multiple templates simultaneously for each data row.
- **Dynamic Placeholder Mapping**: Map any column from your data source to `{{placeholders}}` within your PowerPoint slides.
- **Custom Filename Formatting**: Use placeholders in filenames (e.g., `{{Name}} - {{Date}}.pdf`) to organize output files automatically.
- **Flexible Data Range**: Specify the exact starting and ending rows to process.
- **Output Control**: 
  - Save all documents in a single folder or organize them into separate folders per template.
  - Option to keep generated `.pptx` files alongside the PDFs.
- **Modern Browser Interface**: A sleek, dark-themed GUI built with Streamlit for easy configuration.

## ⚙️ Prerequisites

- **Operating System**: Windows
- **Software**: **Microsoft Office (specifically PowerPoint) must be installed on your system.** This is required as the tool uses PowerPoint's internal engine to ensure high-quality PDF conversions.
- **Python**: Python 3.10+

## 🚀 Installation

1. **Clone or download** this repository to your local machine.
2. **Install dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

## 🛠️ Usage

The easiest way to start the application is by using the provided batch file:

1. Double-click `run_app.bat`.
2. Your browser will automatically open to the GUI (usually at `http://localhost:8501`).

### Detailed Step-by-Step Example

Imagine you want to generate personalized certificates for a course.

**1. Prepare your Data Source**
Create a CSV or Excel file (e.g., `students.csv`) with the following columns:
| Full Name | Course Name | Completion Date |
| :--- | :--- | :--- |
| Alice Smith | Python Basics | 2026-09-01 |
| Bob Jones | Advanced AI | 2026-09-02 |

**2. Prepare your PowerPoint Template**
Create a `.pptx` file and use double curly braces for placeholders where the data should go. For example:
- "This is to certify that **`{{name}}`** has completed the **`{{course}}`** on **`{{date}}`**."

**3. Configure the App**
- **Upload Data**: Upload `students.csv`.
- **Upload Template**: Upload `certificate.pptx`.
- **Map Columns**:
  - `{{name}}` $\rightarrow$ `Full Name`
  - `{{course}}` $\rightarrow$ `Course Name`
  - `{{date}}` $\rightarrow$ `Completion Date`
- **Set Filename Format**: Enter `{{name}} - Certificate`. This will result in files like `Alice Smith - Certificate.pdf`.

**4. Generate**
- Set the **Global Output Directory** (e.g., `C:\Certificates\Output`).
- Choose whether to save in separate folders or one main folder.
- Specify the row range (e.g., Starting Row: 1, Ending Row: 2).
- Click **Start Generation**.

> [!IMPORTANT]
> Because this tool uses COM automation to control Microsoft PowerPoint for PDF conversion, PowerPoint will open and close windows during the generation process. Please do not close PowerPoint manually while the process is running.

## 📁 Project Structure

- `gui_app.py`: The Streamlit web interface.
- `core_logic.py`: The backend engine handling placeholder replacement and PDF conversion.
- `run_app.bat`: One-click launcher for the GUI.
- `requirements.txt`: List of required Python packages.
