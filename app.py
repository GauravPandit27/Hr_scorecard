import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.set_page_config(page_title="Employee Scorecard Generator", layout="wide")

DATA_FILE = "CPF Final Data.xlsx"
TEMPLATE_FILE = "SAMPLE CPF Scorecard- APP Dev-1.xlsx"

def get_col_name(df, keyword):
    """Find a column in df that contains the keyword (case insensitive)."""
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            return col
    return None

def get_col_val(row, keyword):
    """Safely get value from a row based on keyword match for column name."""
    for col in row.index:
        if keyword.lower() in str(col).lower():
            return row[col]
    return ""

@st.cache_data
def load_all_data():
    try:
        xl = pd.ExcelFile(DATA_FILE)
        reporting_sheet = [s for s in xl.sheet_names if "report" in s.lower()][0]
        df_rep = xl.parse(reporting_sheet)
        
        skip_sheets = [s.lower() for s in ['Quick Summary', 'Summary', reporting_sheet]]
        practice_dfs = {}
        for sheet in xl.sheet_names:
            if sheet.lower() not in skip_sheets:
                practice_dfs[sheet.strip()] = xl.parse(sheet)
                
        return df_rep, practice_dfs
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None, None

def match_job_title_column(job_title, practice_df):
    if pd.isnull(job_title):
        return None
    job_title_clean = str(job_title).strip().lower()
    for col in practice_df.columns:
        col_str = str(col).lower()
        if 'capability' in col_str or 'category' in col_str or 'unnamed' in col_str or 'scale' in col_str:
            continue
        
        # Columns often have titles separated by '/'
        col_titles = [t.strip() for t in col_str.split('/')]
        
        # Exact or substring match within the variants
        for t in col_titles:
            if job_title_clean == t or job_title_clean in t:
                return col
    return None

def generate_excel_for_employee(emp_row, edited_df):
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    template_sheet = wb.active
    
    emp_no = get_col_val(emp_row, 'Employee No')
    emp_name = get_col_val(emp_row, 'Employee Name')
    
    safe_name = f"{emp_no}_{emp_name}"[:31]
    safe_name = re.sub(r'[\\/*?:\[\]]', '', safe_name)
    template_sheet.title = safe_name
    
    template_sheet['F3'] = emp_no
    template_sheet['F4'] = emp_name
    template_sheet['F5'] = get_col_val(emp_row, 'Location')
    template_sheet['F6'] = get_col_val(emp_row, 'Department')
    template_sheet['F7'] = get_col_val(emp_row, 'Sub Department')
    template_sheet['F8'] = get_col_val(emp_row, 'Designation')
    template_sheet['F9'] = get_col_val(emp_row, 'Reporting Manager Name')
    
    dj = get_col_val(emp_row, 'Date Joined')
    if pd.notnull(dj):
        template_sheet['F10'] = dj.strftime('%Y-%m-%d') if hasattr(dj, 'strftime') else str(dj)
    else:
        template_sheet['F10'] = ""
        
    template_sheet['F11'] = get_col_val(emp_row, 'Cost Center')
    template_sheet['F12'] = get_col_val(emp_row, 'Primary Practice')
    
    start_row = 17
    
    p_cap_col = get_col_name(edited_df, 'Primary Capability')
    cat_col = get_col_name(edited_df, 'Category')
    cap_col = get_col_name(edited_df, 'Capabilities')
    
    for _, p_row in edited_df.iterrows():
        val_p_cap = p_row[p_cap_col] if (p_cap_col and pd.notnull(p_row[p_cap_col])) else ""
        val_cat = p_row[cat_col] if (cat_col and pd.notnull(p_row[cat_col])) else ""
        val_cap = p_row[cap_col] if (cap_col and pd.notnull(p_row[cap_col])) else ""
        
        cap_cat = p_row.get('Capability Category', "")
        if pd.isnull(cap_cat): cap_cat = ""
        
        target_score = p_row.get('Target Score', "")
        if pd.isnull(target_score): target_score = ""
        
        actual_score = p_row.get('Actual Score', "")
        if pd.isnull(actual_score): actual_score = ""
        
        gap = p_row.get('Gap', "")
        if pd.isnull(gap): gap = ""
        
        def safe_set(coord, val):
            try:
                template_sheet[coord] = val
            except AttributeError:
                pass
                
        safe_set(f'B{start_row}', val_p_cap)
        safe_set(f'C{start_row}', val_cat)
        safe_set(f'D{start_row}', val_cap)
        safe_set(f'E{start_row}', cap_cat)
        safe_set(f'F{start_row}', target_score)
        safe_set(f'G{start_row}', actual_score)
        safe_set(f'H{start_row}', gap)
        
        start_row += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Employee Scorecard Generator")
    
    df_rep, practice_dfs = load_all_data()
    if df_rep is None:
        return
        
    st.sidebar.header("Selections")
    manager_col = get_col_name(df_rep, 'Reporting Manager Name')
    managers = df_rep[manager_col].dropna().unique().tolist()
    managers.sort()
    
    selected_manager = st.sidebar.selectbox("Select Reporting Manager", managers)
    
    if selected_manager:
        emp_df = df_rep[df_rep[manager_col] == selected_manager]
        
        st.sidebar.markdown("---")
        
        emp_no_col = get_col_name(df_rep, 'Employee No')
        emp_name_col = get_col_name(df_rep, 'Employee Name')
        
        emp_list = emp_df.apply(lambda row: f"{row[emp_no_col]} - {row[emp_name_col]}", axis=1).tolist()
        selected_emp_str = st.sidebar.selectbox("Select Employee to View", emp_list)
        
        if selected_emp_str:
            emp_id = int(selected_emp_str.split(" - ")[0])
            emp_row = emp_df[emp_df[emp_no_col] == emp_id].iloc[0]
            
            st.header(f"Scorecard: {emp_row[emp_name_col]}")
            
            st.markdown("### Employee Details")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Employee ID:** {emp_row[emp_no_col]}")
                st.write(f"**Location:** {get_col_val(emp_row, 'Location')}")
                st.write(f"**Department:** {get_col_val(emp_row, 'Department')}")
                st.write(f"**Designation:** {get_col_val(emp_row, 'Designation')}")
                st.write(f"**Cost Center:** {get_col_val(emp_row, 'Cost Center')}")
            with col2:
                st.write(f"**Primary Practice:** {get_col_val(emp_row, 'Primary Practice')}")
                st.write(f"**Sub Department:** {get_col_val(emp_row, 'Sub Department')}")
                st.write(f"**Reporting Manager:** {emp_row[manager_col]}")
                dj = get_col_val(emp_row, 'Date Joined')
                dj_str = dj.strftime('%Y-%m-%d') if pd.notnull(dj) and hasattr(dj, 'strftime') else str(dj)
                st.write(f"**Date Joined:** {dj_str}")
                st.write(f"**Job Title:** {get_col_val(emp_row, 'Job Title')}")
                
            st.markdown("---")
            st.markdown("### Capabilities & Targets")
            
            practice_name = str(get_col_val(emp_row, 'Primary Practice')).strip()
            job_title = get_col_val(emp_row, 'Job Title')
            
            if practice_name in practice_dfs:
                practice_df = practice_dfs[practice_name]
                matched_col = match_job_title_column(job_title, practice_df)
                
                if matched_col:
                    p_cap_col = get_col_name(practice_df, 'Primary Capability')
                    cat_col = None
                    for c in practice_df.columns:
                        cl = str(c).lower()
                        if 'category' in cl and 'capability category' not in cl:
                            cat_col = c
                            break
                    cap_col = get_col_name(practice_df, 'Capabilities')
                    
                    cols_to_extract = []
                    if p_cap_col: cols_to_extract.append(p_cap_col)
                    if cat_col: cols_to_extract.append(cat_col)
                    if cap_col: cols_to_extract.append(cap_col)
                    
                    if cols_to_extract:
                        score_df = practice_df[cols_to_extract + [matched_col]].copy()
                        score_df.rename(columns={matched_col: 'Target Score'}, inplace=True)
                        
                        def format_val(x):
                            if pd.isnull(x): return ''
                            if isinstance(x, float) and x.is_integer(): return str(int(x))
                            return str(x)
                            
                        for col in score_df.columns:
                            score_df[col] = score_df[col].apply(format_val)
                        
                        # Filter out empty capabilities
                        score_df = score_df[score_df[cap_col] != '']
                        
                        state_key = f"df_{emp_id}"
                        editor_key = f"editor_{emp_id}"
                        
                        if state_key not in st.session_state:
                            score_df['Capability Category'] = None
                            score_df['Actual Score'] = None
                            st.session_state[state_key] = score_df.copy()
                            
                        display_df = st.session_state[state_key].copy()
                        
                        if editor_key in st.session_state:
                            edits = st.session_state[editor_key].get("edited_rows", {})
                            for row_idx, row_edits in edits.items():
                                for col, val in row_edits.items():
                                    if col in display_df.columns:
                                        display_df.iloc[int(row_idx), display_df.columns.get_loc(col)] = val
                                        
                        def safe_subtract(target, actual):
                            try:
                                if pd.isnull(actual) or str(actual).strip() == "":
                                    return ""
                                return float(actual) - float(target)
                            except:
                                return ""
                                
                        display_df['Gap'] = display_df.apply(
                            lambda row: safe_subtract(row['Target Score'], row['Actual Score']), axis=1
                        )
                        
                        def color_gap(val):
                            if val == "": return ""
                            try:
                                v = float(val)
                                if v < 0: return "color: red;"
                                if v > 0: return "color: green;"
                                return ""
                            except:
                                return ""
                                
                        styled_df = display_df.style.map(color_gap, subset=['Gap'])
                        
                        with st.form(key=f"form_{emp_id}"):
                            edited_df = st.data_editor(
                                styled_df,
                                key=editor_key,
                                column_config={
                                    "Capability Category": st.column_config.SelectboxColumn(
                                        "Capability Category",
                                        options=["PIC", "BIC", "SC"],
                                    ),
                                    "Actual Score": st.column_config.NumberColumn(
                                        "Actual Score",
                                        min_value=0, max_value=5, step=1,
                                    )
                                },
                                disabled=[p_cap_col, cat_col, cap_col, 'Target Score', 'Gap'],
                                hide_index=True,
                                width='stretch'
                            )
                            submit_button = st.form_submit_button(label='Calculate Gap & Save')
                        
                        st.session_state[state_key]['Actual Score'] = edited_df['Actual Score']
                        st.session_state[state_key]['Capability Category'] = edited_df['Capability Category']
                        
                        st.markdown("---")
                        st.subheader(f"Export Scorecard")
                        
                        if st.button(f"Generate Excel for {emp_row[emp_name_col]}"):
                            with st.spinner("Generating workbook..."):
                                excel_data = generate_excel_for_employee(emp_row, edited_df)
                                st.download_button(
                                    label="Download Scorecard",
                                    data=excel_data,
                                    file_name=f"{emp_row[emp_name_col]}_Scorecard.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                else:
                    st.warning(f"Could not find a matching job title column in '{practice_name}' for '{job_title}'. Ensure the job title exists in the '{practice_name}' sheet headers.")
            else:
                st.warning(f"Practice sheet '{practice_name}' not found in the Data File.")

if __name__ == "__main__":
    main()
