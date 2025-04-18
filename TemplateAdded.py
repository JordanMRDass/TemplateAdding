import streamlit as st
import pandas as pd
import re
import shutil
import xlwings as xw
import tempfile
import os

st.set_page_config(page_title="Excel File Merger", layout="wide")

def number_to_excel_column(n):
    column = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        column = chr(65 + remainder) + column
    return column

def get_source_df(source_file_path, sheet_name = 0):
    source_file = pd.read_excel(source_file_path, sheet_name=sheet_name)

    # Find column header row
    for i, row in enumerate(source_file.values):
        if 'vendor no' in str(row).lower():
            header_row = i
            break

    source_file_df = source_file.copy()
    source_file_df.columns = source_file_df.iloc[header_row]
    source_file_df = source_file_df.iloc[header_row + 1:]

    return source_file_df

def setting_target_configs(target_file_path):
    target_file_df = pd.read_excel(target_file_path, sheet_name="OA")

    last_line = target_file_df.iloc[-1, :].values.tolist()
    target_file_df_columns = target_file_df.columns

    return last_line, target_file_df_columns, target_file_df

def get_new_source_df(target_file_path, source_file_path):
    source_file_df = get_source_df(source_file_path)
    last_line, target_file_df_columns, target_file_df = setting_target_configs(target_file_path)
    
    difference_col = set(target_file_df.columns) - set(source_file_df.columns)
    array = []
    bil_row = int(last_line[0])

    for row in range(len(source_file_df)):
        new_row = []
        for col_num, col in enumerate(target_file_df.columns):
            if col == 'BIL':
                bil_row += 1
                new_row.append(bil_row)

            elif col in difference_col:
                new_row.append("")
            
            else:
                new_row.append(source_file_df[col].values[row])

        array.append(new_row)

    new_source_file_df = pd.DataFrame(array, columns = target_file_df_columns)

    return new_source_file_df, target_file_df, source_file_df

def get_new_path(target_file, source_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tf_target:
        tf_target.write(target_file.read())
        tf_target_path = tf_target.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tf_source:
        tf_source.write(source_file.read())
        tf_source_path = tf_source.name

    new_source_file_df, target_file_df, source_file_df = get_new_source_df(tf_target_path, tf_source_path)
    clean_path = re.sub('\.xlsm','', tf_target_path)
    new_path = f"{clean_path}_new.xlsm"
    shutil.copy(tf_target_path, new_path)

    with col1:
        st.subheader("Target File DataFrame")
        st.write(target_file_df)

    with col2:
        st.subheader("Source File DataFrame")
        st.write(new_source_file_df)

    app = xw.App(visible=False)
    wb = app.books.open(new_path)

    sheet_name = "OA"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        st.error("Cannot find sheet OA")
        return

    last_line_num = len(target_file_df) + 1

    for row in range(len(new_source_file_df)):
        last_line_num += 1
        for col_num, col in enumerate(new_source_file_df.columns):
            column_alpha = number_to_excel_column(col_num + 1)
            sheet.range(f"{column_alpha}{last_line_num}").value = new_source_file_df[col].values[row]

    wb.save()
    wb.close()
    app.quit()

    return new_path, f"{clean_path}_new.xlsm"


# Streamlit App
st.title("Template Adding Tool")

uploaded_target = st.file_uploader("Upload Template File", type="xlsm")
uploaded_source = st.file_uploader("Upload Source File", type="xlsx")

col1, col2 = st.columns(2)

if st.button("Process Files") and uploaded_target and uploaded_source:
    try:
        new_path, new_path_name = get_new_path(uploaded_target, uploaded_source)

        clean_path = re.sub('\.xlsm','', uploaded_target.name)
        new_path_name = f"{clean_path}_new.xlsm"

        with open(new_path, "rb") as f:
            st.download_button(f"Download Processed File: {new_path_name}", f, file_name=new_path_name)
    except Exception as e:
        st.error(f"Error: {e}")
