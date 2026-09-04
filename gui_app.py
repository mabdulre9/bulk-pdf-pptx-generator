import streamlit as st
import pandas as pd
from pathlib import Path
import os
import tempfile
from core_logic import (
    TemplateConfig, 
    extract_placeholders, 
    generate_bulk_documents
)

st.set_page_config(page_title="Bulk PPTX to PDF Generator", layout="wide")

st.title("📄 Bulk PowerPoint to PDF Generator")
st.markdown("Generate personalized PDFs from PowerPoint templates using a data file.")

# --- Session State Initialization ---
if 'templates' not in st.session_state:
    st.session_state.templates = []
if 'df' not in st.session_state:
    st.session_state.df = None

# --- Step 1: Data File ---
st.header("1. Data Source")
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=['csv', 'xlsx', 'xls'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.session_state.df = df
        st.success(f"Loaded {len(df)} rows of data.")
        st.dataframe(df.head(), use_container_width=True)
    except Exception as e:
        st.error(f"Error loading file: {e}")

# --- Step 2: Templates ---
if st.session_state.df is not None:
    st.header("2. PowerPoint Templates")
    
    # Template Upload
    with st.expander("Add PowerPoint Templates", expanded=True):
        template_files = st.file_uploader("Upload PPTX Template(s)", type=['pptx'], accept_multiple_files=True, key="template_uploader")
        if template_files:
            if st.button("Process"):
                for template_file in template_files:
                    # Check if already added to avoid duplicates
                    if any(t.name == template_file.name for t in st.session_state.templates):
                        continue
                    
                    # Save uploaded template to a temp file
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                        tmp.write(template_file.getvalue())
                        tmp_path = Path(tmp.name)
                    
                    placeholders = extract_placeholders(tmp_path)
                    new_template = TemplateConfig(
                        name=template_file.name,
                        path=tmp_path,
                        placeholders=placeholders
                    )
                    st.session_state.templates.append(new_template)
                st.rerun()

    # Template Configuration
    if st.session_state.templates:
        st.subheader("Configure Your Templates")
        for idx, template in enumerate(st.session_state.templates):
            with st.container(border=True):
                st.markdown(f"**Template: {template.name}**")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Column Mapping")
                    if not template.placeholders:
                        st.warning("No placeholders found in this template.")
                    
                    # Create a mapping for this template
                    current_mapping = {}
                    for ph in template.placeholders:
                        selected_col = st.selectbox(
                            f"{{{{{ph}}}}} →", 
                            options=["-- Skip --"] + list(st.session_state.df.columns),
                            key=f"map_{idx}_{ph}"
                        )
                        if selected_col != "-- Skip --":
                            current_mapping[ph] = selected_col
                    
                    template.mapping = current_mapping
                
                with col2:
                    st.markdown("#### Output Settings")
                    
                    fmt_key = f"fmt_{idx}"
                    if fmt_key not in st.session_state:
                        st.session_state[fmt_key] = template.filename_format
                    
                    fmt = st.text_input(
                        "Filename Format", 
                        value=st.session_state[fmt_key],
                        key=f"fmt_input_{idx}",
                        help="Example: {{Name}} - {{Date}} Certificate. Leave empty for default."
                    )
                    st.session_state[fmt_key] = fmt
                    template.filename_format = fmt
                    
                    custom_dir = st.text_input(
                        "Custom Output Folder (Optional)", 
                        value=str(getattr(template, 'custom_output_dir', None)) if getattr(template, 'custom_output_dir', None) else "",
                        key=f"dir_{idx}",
                        placeholder="Leave empty to use global setting",
                        help="Specify a specific folder for documents from this template."
                    )
                    if custom_dir:
                        template.custom_output_dir = Path(custom_dir)
                    else:
                        # Use setattr to avoid AttributeError if object was created with old class def
                        setattr(template, 'custom_output_dir', None)

                    if template.placeholders:
                        st.markdown("**Available Placeholders:**")
                        st.info(" ".join([f"{{{{{ph}}}}}" for ph in template.placeholders]))
                
                if st.button(f"Remove {template.name}", key=f"rem_{idx}"):
                    st.session_state.templates.pop(idx)
                    st.rerun()

    # --- Step 3: Final Execution ---
    if st.session_state.templates:
        st.divider()
        st.header("3. Generate Documents")
        
        with st.form("generation_form"):
            output_path = st.text_input("Global Output Directory Path", placeholder="C:\\Users\\Name\\Documents\\Output")
            separate_folders = st.checkbox("Use separate folders for each template", value=True)
            save_pptx = st.checkbox("Keep generated PPTX files", value=False)
            
            col_start, col_end = st.columns(2)
            with col_start:
                start_row = st.number_input("Starting Row", min_value=1, value=1)
            with col_end:
                end_row = st.number_input("Ending Row", min_value=1, value=len(st.session_state.df) if st.session_state.df is not None else 1)
            
            submit = st.form_submit_button("🚀 Start Generation", use_container_width=True)
            
            if submit:
                if not output_path:
                    st.error("Please provide an output directory path.")
                else:
                    try:
                        out_dir = Path(output_path)
                        out_dir.mkdir(parents=True, exist_ok=True)
                        
                        # Prepare DF using start and end rows
                        working_df = st.session_state.df
                        # Slice dataframe: [start-1 : end]
                        working_df = working_df.iloc[start_row-1 : end_row].reset_index(drop=True)
                        
                        # Progress tracking
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        def progress_callback(current, total, message):
                            percent = current / total
                            progress_bar.progress(percent)
                            status_text.text(f"Processing {current}/{total}: {message}")
                        
                        results = generate_bulk_documents(
                            st.session_state.templates, 
                            working_df, 
                            out_dir, 
                            save_pptx, 
                            separate_folders=separate_folders,
                            progress_callback=progress_callback
                        )
                        
                        st.success("Generation Complete!")
                        
                        # Summary Table
                        summary_data = []
                        for t in results:
                            summary_data.append({
                                "Template": t.name,
                                "Successful (PDF)": t.successful,
                                "PPTX Only": t.pptx_only,
                                "Failed": t.failed
                            })
                        st.table(pd.DataFrame(summary_data))
                        
                    except Exception as e:
                        st.error(f"An error occurred: {e}")

else:
    st.info("Please upload a data file to begin.")
