import streamlit as st
import base64
import PyPDF2
import pandas as pd
import io

st.title("Листы подбора и стикеры для WB")

pdf_file = st.file_uploader("Загрузите PDF файл cо стикерами", type=['pdf'])

xlsx_file = st.file_uploader("Загрузите XLSX файл с информацией о товарах", type=['xlsx'])

if (pdf_file is not None) and (xlsx_file is not None):
    st.write("Вы загрузили файлы:")

    pdf_reader = PyPDF2.PdfReader(pdf_file)
    output_pdf = PyPDF2.PdfWriter()
    ans = []

    for page in pdf_reader.pages:
        text = page.extract_text()
        text = text.split()
        for i in text:
            if i.isdigit():
                pair = (i, page)
                ans.append(pair)

    df = pd.read_excel(xlsx_file)
    first_two_rows = df.head(1).copy()
    selected_columns = first_two_rows.iloc[:, [0, 1, 2, 6]]
    title = selected_columns.columns.values[0]
    data = selected_columns.values[0][0]
    type = selected_columns.values[0][2]
    quantity = selected_columns.values[0][3]

    df = pd.read_excel(xlsx_file, skiprows = 2)
    columns_to_drop = ['Фото', 'Размер', 'Цвет']
    df = df.drop(columns=columns_to_drop)

    value_counts = df['Артикул продавца'].value_counts()
    sorted_df = df.loc[df['Артикул продавца'].isin(value_counts.index)].sort_values(by=['Артикул продавца', 'Бренд'], key=lambda x: x.map(value_counts), ascending=[False, True])
    output_buffer_xlsx = io.BytesIO()
    writer = pd.ExcelWriter(output_buffer_xlsx, engine='xlsxwriter')
    sorted_df.to_excel(writer, sheet_name='Лист подбора', index=False, startrow=2)

    workbook = writer.book
    worksheet = writer.sheets['Лист подбора']

    for idx, col in enumerate(df):
        max_len = max(df[col].astype(str).str.len().max(), len(col))
        worksheet.set_column(idx, idx, max_len + 2)

    for row_num, value in enumerate(sorted_df['Стикер'], start=0):
        worksheet.write_rich_string(row_num+3, 4, value[:-4],  workbook.add_format({'bold': True}), value[-4:]+" ", workbook.add_format({'bold': False}))

    worksheet.merge_range('A1:E1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.merge_range('A2:B2', data)
    worksheet.write('C2', type)
    worksheet.merge_range('D2:E2', quantity)
    workbook.close()

    column_data = sorted_df['Стикер']

    stickers = []

    for i in column_data:
        tmp = i.split(" ")
        tmp = "".join(tmp)
        stickers.append(tmp)

    index_map = {item: index for index, item in enumerate(stickers)}
    sorted_list2 = sorted(ans, key=lambda x: index_map[x[0]])

    for i in sorted_list2:
        output_pdf.add_page(i[1])

    output_buffer_pdf = io.BytesIO()
    output_pdf.write(output_buffer_pdf)
    
    output_buffer_pdf.seek(0)
    output_buffer_xlsx.seek(0)

    pdf_button = st.button("Скачать PDF файл")
    xlsx_button = st.button("Скачать XLSX файл")

    if pdf_button:
        pdf_bytes = output_buffer_pdf.read()
        b64_pdf = base64.b64encode(pdf_bytes).decode()
        href_pdf = f'<a href="data:application/octet-stream;base64,{b64_pdf}" download="downloaded_file.pdf">Нажмите здесь для скачивания PDF</a>'
        st.markdown(href_pdf, unsafe_allow_html=True)

    if xlsx_button:
        xlsx_bytes = output_buffer_xlsx.read()
        b64_xlsx = base64.b64encode(xlsx_bytes).decode()
        href_xlsx = f'<a href="data:application/octet-stream;base64,{b64_xlsx}" download="downloaded_file.xlsx">Нажмите здесь для скачивания XLSX</a>'
        st.markdown(href_xlsx, unsafe_allow_html=True)