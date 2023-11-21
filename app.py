import streamlit as st
import pandas as pd
from datetime import datetime
import os
import base64
from io import BytesIO

def read_data(file):
    return pd.read_excel(file, na_values=['NA'])

def find_differences(df1, df2):
    merged_data = df1.merge(df2, left_on='ID', right_on='ID', how='outer')
    return merged_data

def format_excel(writer, sheet_name, df):
    df.to_excel(writer, sheet_name=sheet_name, startrow=3)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Add title
    worksheet.write(0, 0, 'Monitoring tool ' + datetime.now().strftime('%d %b %Y'),
                    workbook.add_format({'bold': True, 'color': 'red', 'size': 16}))

    # Header formatting
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'fg_color': '#FDE9D9', 'border': 1})
    for col_num, value in enumerate(df.columns.values, 1):
        worksheet.write(3, col_num, value, header_format)

    # Color yellow for changed cells
    changed_cell = workbook.add_format({'bg_color': 'yellow'})
    worksheet.conditional_format('A1:BH10000', {'type': 'text', 'criteria': 'containing', 'value': '-->',
                                                 'format': changed_cell})

    # Apply filter
    worksheet.autofilter('A4:BO10000')

def save_to_excel(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        format_excel(writer, sheet_name, df)

    return output

def main():
    st.title('Excel Comparison Tool')

    st.sidebar.header('Select your 2 Excel files')
    uploaded_file1 = st.sidebar.file_uploader("Upload the first Excel file here", type=["xlsx"])
    uploaded_file2 = st.sidebar.file_uploader("Upload the second Excel file here", type=["xlsx"])

    if st.sidebar.button("Process"):
        if uploaded_file1 is not None and uploaded_file2 is not None:
            df1 = pd.read_excel(uploaded_file1)
            df2 = pd.read_excel(uploaded_file2)

            merged_data = find_differences(df1, df2)

            df1 = df1.set_index('ID')
            df2 = df2.set_index('ID')
            df3 = pd.concat([df1, df2], sort=False)
            df3a = df3.stack().explode().astype(str).groupby(level=[0, 1]).unique().transform(lambda x: '--> '.join(x)).unstack(1)

            df3a.loc[~df3a.index.isin(df2.index), 'status'] = 'deleted'
            df3a.loc[~df3a.index.isin(df1.index), 'status'] = 'new'
            idx = df3.stack().groupby(level=[0, 1]).nunique()
            df3a.loc[idx.mask(idx <= 1).dropna().index.get_level_values(0), 'status'] = 'modified'

            df3b = df3a[df3a.status != 'nan']
            df3b.insert(0, 'status', df3b.pop('status'))

            if df3b.empty:
                st.warning("No data to display.")
            else:
                st.write(df3b)

                # Provide download link
                st.markdown(get_download_link(df3b, "Monitoring"), unsafe_allow_html=True)

def get_download_link(df, sheet_name):
    """Generate a download link for the DataFrame."""
    output = save_to_excel(df, sheet_name)
    base64_encoded = base64.b64encode(output.getvalue()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64_encoded}" download="output.xlsx">Download Excel File</a>'

if __name__ == '__main__':
    main()
