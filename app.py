import streamlit as st
import pandas as pd
import base64
from io import BytesIO

def get_excel_sheets(file):
    """Get list of sheet names from Excel file"""
    try:
        excel_file = pd.ExcelFile(file)
        return excel_file.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return []

def find_and_read_data(file_path, headers, sheet_name=0):
    df = pd.read_excel(file_path, header=None, sheet_name=sheet_name)
    header_row = None

    # Locate the header row
    for i, row in df.iterrows():
        if set(headers).issubset(set(row)):
            header_row = i
            break

    if header_row is None:
        raise ValueError(f"Headers {headers} not found in sheet '{sheet_name}' of the Excel file")

    df.columns = df.iloc[header_row]
    df = df[header_row+1:]
    return df[headers]

def detect_instansi_column(uploaded_files, selected_sheets):
    """Detect if INSTANSI column exists in uploaded files"""
    for i, uploaded_file in enumerate(uploaded_files):
        try:
            sheet_name = selected_sheets.get(uploaded_file.name, 0)
            # Read first few rows to check headers
            df = pd.read_excel(uploaded_file, header=None, nrows=10, sheet_name=sheet_name)
            for j, row in df.iterrows():
                if 'INSTANSI' in row.values:
                    return True
        except:
            continue
    return False

def get_missing_ref_instansi(missing_instansi_list):
    """Get manual input for missing REF_INSTANSI values"""
    manual_mapping = {}
    
    if missing_instansi_list:
        st.subheader("🔧 Manual REF_INSTANSI Input")
        st.write("The following INSTANSI values were not found in the reference file. Please provide REF_INSTANSI codes:")
        
        for instansi in missing_instansi_list:
            ref_code = st.text_input(
                f"REF_INSTANSI for '{instansi}':",
                key=f"ref_{instansi}",
                help=f"Enter the reference code for {instansi}"
            )
            if ref_code:
                manual_mapping[instansi] = ref_code
    
    return manual_mapping

def expand_instansi_rows(df):
    """Expand rows that contain multiple INSTANSI values separated by comma"""
    expanded_rows = []
    
    for _, row in df.iterrows():
        instansi_value = str(row['INSTANSI']).strip()
        
        # Check if INSTANSI contains multiple values (separated by comma, semicolon, or pipe)
        if ',' in instansi_value or ';' in instansi_value or '|' in instansi_value:
            # Split by multiple possible separators
            instansi_list = []
            for sep in [',', ';', '|']:
                if sep in instansi_value:
                    instansi_list = [inst.strip() for inst in instansi_value.split(sep) if inst.strip()]
                    break
            
            # Create a row for each INSTANSI
            for instansi in instansi_list:
                new_row = row.copy()
                new_row['INSTANSI'] = instansi
                expanded_rows.append(new_row)
        else:
            # Single INSTANSI value
            expanded_rows.append(row)
    
    return pd.DataFrame(expanded_rows)

def get_ref_instansi_for_value(instansi_value, ref_instansi_dict=None, manual_mapping=None):
    """Get REF_INSTANSI for a given INSTANSI value"""
    # First check manual mapping
    if manual_mapping and instansi_value in manual_mapping:
        return manual_mapping[instansi_value]
    
    # Then check reference dict
    if ref_instansi_dict:
        for name, code in ref_instansi_dict.items():
            if code == instansi_value or name.upper() == instansi_value.upper():
                return code
    
    # Return None if not found
    return None

def find_missing_instansi(dataframes, ref_instansi_dict=None):
    """Find INSTANSI values that don't have REF_INSTANSI mapping"""
    all_instansi = set()
    
    for df in dataframes:
        if 'INSTANSI' in df.columns:
            expanded_df = expand_instansi_rows(df)
            expanded_df['INSTANSI'] = expanded_df['INSTANSI'].astype(str).str.strip()
            expanded_df = expanded_df[expanded_df['INSTANSI'].notna()]
            expanded_df = expanded_df[expanded_df['INSTANSI'] != '']
            expanded_df = expanded_df[expanded_df['INSTANSI'] != 'nan']
            all_instansi.update(expanded_df['INSTANSI'].unique())
    
    missing_instansi = []
    for instansi in all_instansi:
        found = False
        if ref_instansi_dict:
            for name, code in ref_instansi_dict.items():
                if code == instansi or name.upper() == instansi.upper():
                    found = True
                    break
        if not found:
            missing_instansi.append(instansi)
    
    return sorted(missing_instansi)

def separate_data(dataframes, has_instansi_column, selected_instansi_name=None, ref_instansi_code=None, ref_instansi_dict=None, manual_mapping=None):
    """Separate data automatically based on available columns"""
    separated_dfs = {}
    
    for df in dataframes:
        if has_instansi_column:
            # Mode: JENIS TES + INSTANSI
            # First, expand rows with multiple INSTANSI values
            expanded_df = expand_instansi_rows(df)
            
            # Clean and validate INSTANSI values
            expanded_df['INSTANSI'] = expanded_df['INSTANSI'].astype(str).str.strip()
            expanded_df = expanded_df[expanded_df['INSTANSI'].notna()]
            expanded_df = expanded_df[expanded_df['INSTANSI'] != '']
            expanded_df = expanded_df[expanded_df['INSTANSI'] != 'nan']
            
            if expanded_df.empty:
                continue
            
            # Add REF_INSTANSI column based on mapping
            expanded_df['REF_INSTANSI'] = expanded_df['INSTANSI'].apply(
                lambda x: get_ref_instansi_for_value(x, ref_instansi_dict, manual_mapping)
            )
                
            # Group by JENIS TES and INSTANSI
            try:
                for (jenis_tes, instansi), group_df in expanded_df.groupby(['JENIS TES', 'INSTANSI']):
                    key = f"{instansi}_{jenis_tes}"
                    
                    if key not in separated_dfs:
                        separated_dfs[key] = []
                    separated_dfs[key].append(group_df)
                    
            except Exception as e:
                st.error(f"Error processing dataframe: {str(e)}")
                continue
        else:
            # Mode: JENIS TES only (but add INSTANSI and REF_INSTANSI columns)
            # Add INSTANSI and REF_INSTANSI columns for consistency
            df['INSTANSI'] = selected_instansi_name or "NO_INSTANSI"
            df['REF_INSTANSI'] = ref_instansi_code
            
            for jenis_tes, group_df in df.groupby('JENIS TES'):
                key = f"{selected_instansi_name or 'NO_INSTANSI'}_{jenis_tes}"
                
                if key not in separated_dfs:
                    separated_dfs[key] = []
                separated_dfs[key].append(group_df)
    
    # Combine dataframes for each combination
    final_separated = {}
    for key, df_list in separated_dfs.items():
        final_separated[key] = pd.concat(df_list, ignore_index=True)
    
    return final_separated

def create_download_link(df, filename, link_text):
    """Create download link for Excel file"""
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine='xlsxwriter')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:file/xlsx;base64,{b64}" download="{filename}" style="text-decoration: none; background-color: #4CAF50; color: white; padding: 8px 16px; border-radius: 4px; display: inline-block; margin: 5px;">{link_text}</a>'
    return href

def main():
    st.set_page_config(page_title="Data Peserta Generator", page_icon="📊", layout="wide")
    
    st.title('📊 Data Peserta Generator')
    st.markdown("---")

    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Load the predefined ref_instansi list from an Excel file
        st.subheader("📁 Reference Instansi")
        st.caption("🔹 Optional: Upload untuk mapping REF_INSTANSI otomatis")
        
        ref_instansi_file = st.file_uploader(
            "Upload Reference Instansi Excel file (Optional)", 
            type=['xlsx'], 
            help="Upload file containing INSTANSI and REF_INSTANSI columns for automatic REF_INSTANSI mapping"
        )
        
        ref_instansi_dict = {}
        instansi_options = ["None", "Custom"]
        
        if ref_instansi_file is not None:
            try:
                ref_instansi_df = pd.read_excel(ref_instansi_file)
                ref_instansi_dict = dict(zip(ref_instansi_df['INSTANSI'], ref_instansi_df['REF_INSTANSI']))
                instansi_options.extend(list(ref_instansi_dict.keys()))
                st.success(f"✅ Loaded {len(ref_instansi_dict)} instansi references")
                
                # Show preview of reference data
                with st.expander("Preview Reference Data"):
                    st.dataframe(ref_instansi_df.head(), use_container_width=True)
                    
            except Exception as e:
                st.error(f"❌ Error reading reference file: {str(e)}")

        st.markdown("---")
        
        # New location input
        new_lokasi_ujian = st.text_input(
            "🏢 Masukkan bila ingin merubah LOKASI_UJIAN:", 
            "",
            help="Leave empty to keep original location from files"
        )

    # Main content area
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📂 Upload Files")
        uploaded_files = st.file_uploader(
            "Upload Excel files", 
            type=['xlsx'], 
            accept_multiple_files=True,
            help="Select multiple Excel files containing participant data"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name}")

    with col2:
        st.subheader("📋 File Format Requirements")
        
        st.write("**Required columns:**")
        st.write("- NO_PESERTA")
        st.write("- NAMA")
        st.write("- JENIS TES") 
        st.write("- LOKASI UJIAN")
        st.write("- INSTANSI (optional - for automatic separation)")
        
        st.write("**Output akan include:** INSTANSI & REF_INSTANSI")
        st.caption("💡 Jika ada kolom INSTANSI: dipisahkan berdasarkan INSTANSI & JENIS TES")
        st.caption("💡 Jika tidak ada kolom INSTANSI: dipisahkan berdasarkan JENIS TES saja")
        st.caption("💡 INSTANSI dapat berisi beberapa nilai dipisahkan dengan koma (,), titik koma (;), atau pipe (|)")

    st.markdown("---")

    if uploaded_files:
        # Sheet selection section
        st.subheader("📋 Sheet Selection")
        st.write("Select the sheet to process for each Excel file:")
        
        selected_sheets = {}
        sheet_selection_complete = True
        
        for uploaded_file in uploaded_files:
            sheets = get_excel_sheets(uploaded_file)
            if sheets:
                col1, col2 = st.columns([1, 2])
                with col1:
                    st.write(f"**{uploaded_file.name}**")
                with col2:
                    selected_sheet = st.selectbox(
                        f"Select sheet for {uploaded_file.name}:",
                        options=sheets,
                        key=f"sheet_{uploaded_file.name}",
                        help=f"Available sheets: {', '.join(sheets)}"
                    )
                    selected_sheets[uploaded_file.name] = selected_sheet
                    
                    # Show preview of selected sheet
                    with st.expander(f"Preview {selected_sheet}", expanded=False):
                        try:
                            preview_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=5)
                            st.dataframe(preview_df, use_container_width=True)
                        except Exception as e:
                            st.error(f"Error previewing sheet: {str(e)}")
            else:
                st.error(f"❌ Could not read sheets from {uploaded_file.name}")
                sheet_selection_complete = False
        
        if not sheet_selection_complete:
            st.stop()
        
        st.markdown("---")
        
        # Detect if INSTANSI column exists
        has_instansi_column = detect_instansi_column(uploaded_files, selected_sheets)
        
        if has_instansi_column:
            st.info("🔍 Detected INSTANSI column - will separate by INSTANSI & JENIS TES")
            headers = ['NO_PESERTA', 'NAMA', 'JENIS TES', 'LOKASI UJIAN', 'INSTANSI']
        else:
            st.info("🔍 No INSTANSI column detected - will separate by JENIS TES only")
            headers = ['NO_PESERTA', 'NAMA', 'JENIS TES', 'LOKASI UJIAN']
            
            # Show instansi selection for JENIS TES only mode
            st.subheader("🏢 Instansi Configuration")
            selected_instansi = st.selectbox(
                "Select a reference instansi:", 
                instansi_options,
                help="Choose from predefined instansi or select 'Custom' for manual input"
            )

            ref_instansi_code = None
            selected_instansi_name = None

            if selected_instansi == "Custom":
                selected_instansi_name = st.text_input(
                    "Enter instansi name:",
                    help="Enter instansi name"
                )
                ref_instansi_code = st.text_input(
                    "Enter REF_INSTANSI code:",
                    help="Enter your REF_INSTANSI reference code"
                )
                selected_instansi_name = selected_instansi_name if selected_instansi_name else "CUSTOM"
            elif selected_instansi != "None":
                ref_instansi_code = ref_instansi_dict.get(selected_instansi)
                selected_instansi_name = selected_instansi
                if ref_instansi_code:
                    st.info(f"📋 Selected: {selected_instansi_name} -> {ref_instansi_code}")
            
        try:
            with st.spinner("🔄 Processing files..."):
                dataframes = []
                
                for uploaded_file in uploaded_files:
                    sheet_name = selected_sheets[uploaded_file.name]
                    st.write(f"Processing {uploaded_file.name} - Sheet: {sheet_name}")
                    
                    df = find_and_read_data(uploaded_file, headers, sheet_name)
                    
                    # Rename columns
                    rename_dict = {
                        'NO_PESERTA': 'PARTICIPANT_NO',
                        'NAMA': 'NAME',
                        'JENIS TES': 'JENIS TES',
                        'LOKASI UJIAN': 'LOKASI'
                    }
                    
                    if has_instansi_column:
                        rename_dict['INSTANSI'] = 'INSTANSI'
                    
                    df.rename(columns=rename_dict, inplace=True)
                    
                    # Add NIK column
                    df['NIK'] = df['PARTICIPANT_NO']

                    # Update 'LOKASI_UJIAN' bila ada inputan 
                    if new_lokasi_ujian:
                        df['LOKASI'] = new_lokasi_ujian

                    dataframes.append(df)

            # Handle missing REF_INSTANSI mappings for INSTANSI mode
            manual_mapping = {}
            if has_instansi_column:
                missing_instansi = find_missing_instansi(dataframes, ref_instansi_dict)
                if missing_instansi:
                    manual_mapping = get_missing_ref_instansi(missing_instansi)
                    
                    # Check if all missing instansi have been provided with REF_INSTANSI
                    missing_refs = [inst for inst in missing_instansi if inst not in manual_mapping or not manual_mapping[inst]]
                    if missing_refs:
                        st.warning(f"⚠️ Please provide REF_INSTANSI for: {', '.join(missing_refs)}")
                        st.stop()

            # Separate data
            if has_instansi_column:
                separated_data = separate_data(
                    dataframes, 
                    has_instansi_column=True, 
                    ref_instansi_dict=ref_instansi_dict, 
                    manual_mapping=manual_mapping
                )
                st.subheader("📊 Data separated by INSTANSI & JENIS TES")
            else:
                separated_data = separate_data(
                    dataframes, 
                    has_instansi_column=False, 
                    selected_instansi_name=selected_instansi_name, 
                    ref_instansi_code=ref_instansi_code
                )
                st.subheader("📊 Data separated by JENIS TES")

            # Display results
            if separated_data:
                # Summary statistics
                total_records = sum(len(df) for df in separated_data.values())
                st.metric("Total Records Processed", total_records)
                
                # Show expansion info for INSTANSI mode
                if has_instansi_column:
                    original_records = sum(len(df) for df in dataframes)
                    if total_records > original_records:
                        st.info(f"ℹ️ {total_records - original_records} additional records created from expanding multiple INSTANSI values")
                
                # Create tabs for better organization
                tab1, tab2 = st.tabs(["📋 Data Preview", "⬇️ Download Files"])
                
                with tab1:
                    for key, combined_df in separated_data.items():
                        with st.expander(f"📄 {key} ({len(combined_df)} records)", expanded=False):
                            st.dataframe(combined_df.head(10), use_container_width=True)
                            if len(combined_df) > 10:
                                st.caption(f"Showing first 10 of {len(combined_df)} records")

                with tab2:
                    st.write("Click the buttons below to download Excel files:")
                    
                    # Create download links
                    download_links = []
                    for key, combined_df in separated_data.items():
                        filename = f"data_{key}.xlsx"
                        
                        link = create_download_link(
                            combined_df, 
                            filename, 
                            f"📥 Download {key} ({len(combined_df)} records)"
                        )
                        download_links.append(link)
                    
                    # Display download links in columns
                    cols = st.columns(2)
                    for i, link in enumerate(download_links):
                        with cols[i % 2]:
                            st.markdown(link, unsafe_allow_html=True)
                            st.write("")  # Add spacing

                # Additional statistics
                st.markdown("---")
                st.subheader("📈 Statistics")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Files Generated", len(separated_data))
                with col2:
                    st.metric("Total Records", total_records)
                with col3:
                    avg_records = total_records / len(separated_data) if separated_data else 0
                    st.metric("Avg Records per File", f"{avg_records:.1f}")

                # Show INSTANSI and REF_INSTANSI mapping
                st.markdown("---")
                st.subheader("📋 INSTANSI & REF_INSTANSI Mapping")
                
                # Collect all INSTANSI and their REF_INSTANSI from the processed data
                mapping_data = []
                for key, df in separated_data.items():
                    if not df.empty and 'INSTANSI' in df.columns and 'REF_INSTANSI' in df.columns:
                        unique_mappings = df[['INSTANSI', 'REF_INSTANSI']].drop_duplicates()
                        mapping_data.extend(unique_mappings.to_dict('records'))
                
                if mapping_data:
                    mapping_df = pd.DataFrame(mapping_data).drop_duplicates()
                    st.dataframe(mapping_df, use_container_width=True)

                # Show processed files and sheets
                st.markdown("---")
                st.subheader("📁 Processed Files & Sheets")
                processed_info = []
                for file_name, sheet_name in selected_sheets.items():
                    processed_info.append({"File": file_name, "Sheet": sheet_name})
                
                if processed_info:
                    processed_df = pd.DataFrame(processed_info)
                    st.dataframe(processed_df, use_container_width=True)

            else:
                st.warning("⚠️ No data found after processing files.")

        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.write("Please check:")
            st.write("- File format is correct (.xlsx)")
            st.write("- Selected sheets contain the required headers:", headers)
            st.write("- Files are not corrupted")
            if has_instansi_column:
                st.write("- INSTANSI column contains valid values")

    else:
        st.info("👆 Please upload Excel files to begin processing")
        
        # Show example of expected file format
        with st.expander("📋 Expected File Format Examples", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**With INSTANSI column (auto-separate by INSTANSI & JENIS TES):**")
                example_df1 = pd.DataFrame({
                    'NO_PESERTA': ['001', '002', '003'],
                    'NAMA': ['John Doe', 'Jane Smith', 'Bob Johnson'],
                    'JENIS TES': ['CPNS', 'PPPK', 'CPNS'],
                    'LOKASI UJIAN': ['Jakarta', 'Bandung', 'Surabaya'],
                    'INSTANSI': ['KEMKES', 'POLRI,KEMKES', 'KEMKES']
                })
                st.dataframe(example_df1, use_container_width=True)
            
            with col2:
                st.write("**Without INSTANSI column (separate by JENIS TES only):**")
                example_df2 = pd.DataFrame({
                    'NO_PESERTA': ['001', '002', '003'],
                    'NAMA': ['John Doe', 'Jane Smith', 'Bob Johnson'],
                    'JENIS TES': ['CPNS', 'PPPK', 'CPNS'],
                    'LOKASI UJIAN': ['Jakarta', 'Bandung', 'Surabaya']
                })
                st.dataframe(example_df2, use_container_width=True)

if __name__ == "__main__":
    main()
