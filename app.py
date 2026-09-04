import streamlit as st
import zipfile
import pandas as pd
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Zip CSV Comparator", layout="centered")

st.title("📊 CSV Zip Comparator")
st.write("Drag and drop two zip files below. The app will pair up the CSVs inside them sequentially and generate an Excel report of the mismatched rows.")

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    zip1_file = st.file_uploader("Upload First Zip (e.g., UATDT)", type=["zip"])
with col2:
    zip2_file = st.file_uploader("Upload Second Zip (e.g., STGDT)", type=["zip"])

# --- PROCESSING LOGIC ---
if zip1_file and zip2_file:
    if st.button("Run Comparison"):
        with st.spinner("Analyzing and comparing rows..."):
            try:
                # Set up an in-memory buffer to hold the final Excel file
                output_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip1_file, 'r') as z1, zipfile.ZipFile(zip2_file, 'r') as z2:
                    z1_csvs = sorted([f for f in z1.namelist() if f.lower().endswith('.csv')])
                    z2_csvs = sorted([f for f in z2.namelist() if f.lower().endswith('.csv')])
                    
                    if not z1_csvs or not z2_csvs:
                        st.error("Could not find CSV files in one or both zip archives.")
                    else:
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            for i, (f1_name, f2_name) in enumerate(zip(z1_csvs, z2_csvs)):
                                
                                with z1.open(f1_name) as f1, z2.open(f2_name) as f2:
                                    df1 = pd.read_csv(io.BytesIO(f1.read()))
                                    df2 = pd.read_csv(io.BytesIO(f2.read()))
                                    
                                    comparison = pd.merge(
                                        df1.astype(str), 
                                        df2.astype(str), 
                                        how='outer', 
                                        indicator='Difference_Location'
                                    )
                                    
                                    diff_only = comparison[comparison['Difference_Location'] != 'both'].copy()
                                    sheet_name = f"CSV_Pair_{i+1}"
                                    
                                    if not diff_only.empty:
                                        diff_only['Difference_Location'] = diff_only['Difference_Location'].map({
                                            'left_only': 'Only in Zip 1',
                                            'right_only': 'Only in Zip 2'
                                        })
                                        diff_only.to_excel(writer, sheet_name=sheet_name, index=False)
                                    else:
                                        pd.DataFrame({'Result': ['No differences found!']}).to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Ready the buffer for downloading
                output_buffer.seek(0)
                
                st.success("Comparison complete! Your report is ready.")
                
                # Create the download button
                st.download_button(
                    label="📥 Download Excel Report",
                    data=output_buffer,
                    file_name="differences_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")