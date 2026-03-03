import streamlit as st
import pandas as pd
from io import BytesIO

# ── Hardcoded client-to-company mapping (Client Ref No → PPE or PAM) ──
CLIENT_MAPPING = {
    # PW8 (CC- codes)
    'CC-4429': 'PPE',
    'CC-4819': 'PPE',
    'CC-4807': 'PPE',
    'CC-4735': 'PPE',
    'CC-4794': 'PPE',
    'CC-4432': 'PPE',
    'CC-4504': 'PPE',
    'CC-4752': 'PPE',
    'CC-4386': 'PPE',
    'CC-4861': 'PPE',
    'CC-4499': 'PAM',
    'CC-4430': 'PAM',
    'CC-4581': 'PAM',
    'CC-4653': 'PAM',
    'CC-4580': 'PAM',
    'CC-4582': 'PAM',
    # PW9 (IC- codes)
    'IC-726952': 'PPE',
    'IC-556522': 'PPE',
    'IC-785908': 'PPE',
    'IC-739860': 'PPE',
    'IC-758379': 'PPE',
    'IC-562665': 'PPE',
    'IC-727503': 'PPE',
    'IC-751589': 'PPE',
    'IC-746761': 'PPE',
    'IC-727722': 'PPE',
    'IC-727564': 'PPE',
    'IC-786821': 'PPE',
    'IC-727723': 'PPE',
    'IC-751615': 'PPE',
    'IC-750672': 'PPE',
    'IC-752165': 'PPE',
    'IC-762480': 'PAM',
    'IC-749310': 'PAM',
    'IC-747918': 'PAM',
    'IC-747040': 'PAM',
    'IC-745967': 'PAM',
    'IC-725861': 'PAM',
    'IC-740454': 'PAM',
    'IC-745103': 'PAM',

    # PM (IC- codes)
    'IC-831931': 'PM',  
    'IC-831937': 'PM', 
    'IC-831934': 'PM',  
    'IC-831906': 'PM',  
    'IC-831935': 'PM',  
    'IC-831919': 'PM',
    'IC-831913': 'PM',  
    'IC-831917': 'PM',  
    'IC-831911': 'PM',  
    'IC-831908': 'PM',  
    'IC-831918': 'PM',  
    'IC-831932': 'PM',  
    'IC-831909': 'PM', 
    'IC-831925': 'PM',
    'IC-831929': 'PM',
    'IC-831944': 'PM', 
    'IC-831938': 'PM', 
}

# Columns to include in output (all others will be excluded)
INCLUDE_COLS = ['Client Name', 'Portfolio Value']

def to_excel_bytes(df):
    """Convert DataFrame to Excel bytes for download."""
    buf = BytesIO()
    
    # Create Excel writer object
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Apply left alignment to all cells
        from openpyxl.styles import Alignment
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='left')
        
        # Apply currency format to Portfolio Value column if it exists
        if 'Portfolio Value' in df.columns:
            col_idx = df.columns.get_loc('Portfolio Value') + 1  # +1 because Excel is 1-indexed
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            
            # Apply currency format to all cells in Portfolio Value column (except header)
            for row_num in range(2, len(df) + 2):  # Start from row 2 to skip header
                cell = worksheet[f'{col_letter}{row_num}']
                cell.number_format = '$#,##0.00'
    
    return buf.getvalue()

def process_file(uploaded_file, code_label):
    """Read uploaded Excel, classify rows, return PPE/PAM/Unknown DataFrames."""
    df = pd.read_excel(uploaded_file)

    # Drop the Referral column if it exists (since it was manually added)
    if 'Referral' in df.columns or 'Referral ' in df.columns:
        df = df.drop(columns=[c for c in df.columns if c.strip() == 'Referral'])

    # Identify the Client Ref No column (first column)
    ref_col = df.columns[0]

    # Classify each row
    df['_company'] = df[ref_col].map(CLIENT_MAPPING).fillna('UNKNOWN')
    df['_source'] = code_label

    ppe = df[df['_company'] == 'PPE'].drop(columns=['_company', '_source'])
    pam = df[df['_company'] == 'PAM'].drop(columns=['_company', '_source'])
    pm = df[df['_company'] == 'PM'].drop(columns=['_company', '_source'])
    unknown = df[df['_company'] == 'UNKNOWN'].drop(columns=['_company', '_source'])

    return ppe, pam, pm, unknown


def filter_output_cols(df):
    """Keep only Client Name and Portfolio Value columns in output, with sequential numbering."""
    # Get columns to keep (only those in INCLUDE_COLS that exist in the dataframe)
    cols_to_keep = [c for c in df.columns if c in INCLUDE_COLS]
    
    # Create a new dataframe with sequential numbering
    result_df = pd.DataFrame()
    result_df['No.'] = range(1, len(df) + 1)
    
    # Add the columns we want to keep
    for col in cols_to_keep:
        if col in df.columns:
            result_df[col] = df[col].values
    
    return result_df


# ─────────────────────── Streamlit UI ───────────────────────
st.set_page_config(page_title="PW Client Sorter", page_icon="📊", layout="wide")

st.title("📊 PW8 / PW9 Client Sorter")
st.markdown("Upload your **PW8** and **PW9** client Excel files. "
            "Clients will be sorted into **PPE** (Phillip Private Equity) and **PAM** (PAM Australia) "
            "based on internal mapping. Unmatched clients go to **Unknown**.")

st.divider()

col1, col2 = st.columns(2)
with col1:
    pw8_file = st.file_uploader("Upload PW8 Client Excel", type=["xlsx", "xls"], key="pw8")
with col2:
    pw9_file = st.file_uploader("Upload PW9 Client Excel", type=["xlsx", "xls"], key="pw9")

if pw8_file and pw9_file:
    with st.spinner("Processing files..."):
        pw8_ppe, pw8_pam, pw8_pm, pw8_unk = process_file(pw8_file, 'PW8')
        pw9_ppe, pw9_pam, pw9_pm, pw9_unk = process_file(pw9_file, 'PW9')

        # Combine PW8 + PW9 results
        all_ppe = pd.concat([pw8_ppe, pw9_ppe], ignore_index=True)
        all_pam = pd.concat([pw8_pam, pw9_pam], ignore_index=True)
        all_pm = pd.concat([pw8_pm, pw9_pm], ignore_index=True)
        all_unknown = pd.concat([pw8_unk, pw9_unk], ignore_index=True)

        # Keep only Client Name and Portfolio Value for PPE & PAM outputs
        ppe_output = filter_output_cols(all_ppe)
        pam_output = filter_output_cols(all_pam)
        pm_output = filter_output_cols(all_pm)
        unk_output = all_unknown  # Keep all columns for unknown so you can identify them

    st.success(f"✅ Done! — **PPE**: {len(ppe_output)} clients · **PAM**: {len(pam_output)} clients · **Unknown**: {len(unk_output)} clients")

    st.divider()

    # ── PPE Output ──
    tab1, tab2, tab3 ,tab4 = st.tabs(["🔵 PPE (Phillip Private Equity)", "🟢 PAM (Australia)", "🔴 Unknown", "🟣 PM (Pinnacle Marine)"])

    with tab1:
        st.subheader(f"PPE Clients ({len(ppe_output)})")
        st.dataframe(ppe_output, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download PPE Excel",
            data=to_excel_bytes(ppe_output),
            file_name="PPE_Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab2:
        st.subheader(f"PAM Clients ({len(pam_output)})")
        st.dataframe(pam_output, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download PAM Excel",
            data=to_excel_bytes(pam_output),
            file_name="PAM_Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab3:
        st.subheader(f"Unknown Clients ({len(unk_output)})")
        if len(unk_output) > 0:
            st.warning("These clients were not found in either PPE or PAM mapping.")
            st.dataframe(unk_output, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download Unknown Excel",
                data=to_excel_bytes(unk_output),
                file_name="Unknown_Clients.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("🎉 All clients matched — no unknowns!")

    with tab4:
        st.subheader(f"Pinnacle Marine Clients ({len(pm_output)})")
        st.dataframe(pm_output, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download PM Excel",
            data=to_excel_bytes(pm_output),
            file_name="PinnacleMarine_Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif pw8_file or pw9_file:
    st.info("Please upload **both** PW8 and PW9 files to proceed.")