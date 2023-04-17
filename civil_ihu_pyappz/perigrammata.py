import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
# from docxcompose.composer import Composer
# from docx import Document
# from docx.enum.text import WD_BREAK
from pathlib import Path
import json

import io
import requests



def replace_none_with_empty_str(some_dict: dict) -> dict:
    return {k: ('' if v is None else v) for k, v in some_dict.items()}


try:
    with open('keys.json') as f:
       data = json.load(f)
       gsheet_perigrammata_id = data['gsheet_perigrammata']
except:
    gsheet_perigrammata_id = st.secrets['gsheet_perigrammata_id']



st.markdown('## Περιγράμματα μαθημάτων')

url = r"https://github.com/panagop/civil_ihu_pyappz/raw/daf2ec082269559dcc28c8675b0a55178fd3b122/civil_ihu_pyappz/Perigrammata-template-gr.docx"
response = requests.get(url)
bytes_io = io.BytesIO(response.content)

doc = DocxTemplate(bytes_io)


# doc = DocxTemplate("Perigrammata-template-gr.docx")
# doc_examino = DocxTemplate("Εξάμηνο-template-gr.docx")

st.write(doc.undeclared_template_variables)

lang = st.radio("Γλώσσα", ("Ελληνικά", "Αγγλικά"))


@st.cache_data
def load_gheet(lang):
    sheet_id = gsheet_perigrammata_id
    if lang == "Ελληνικά":
        sheet_name = "gr"
    else:
        sheet_name = "eng"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url, dtype_backend='pyarrow', index_col=0)
    return df


df = load_gheet(lang)



tab_table, tab_statistics, tab_download = st.tabs(
    ["Πίνακας", "Στατιστικά", "Αρχείο word"])

with tab_table:
    st.write(df)

with tab_statistics:
    st.markdown('### Αριθμός μαθημάτων ανά εξάμηνο')
    st.bar_chart(df['examino'].value_counts())

    st.markdown('### Τύπος μαθημάτων')
    st.bar_chart(df['type'].value_counts())

with tab_download:

    docx_examino = st.selectbox("Επιλέξτε εξάμηνο", df['examino'].unique())
    docx_code = st.selectbox("Επιλέξτε κωδικό μαθήματος", df[df['examino'] == docx_examino]['code'].unique())

    row_index = df[(df['examino'] == docx_examino) & (df['code'] == docx_code)].index[0]-1


    row = df.iloc[row_index]
    row_dict = row.to_dict()
    row_dict = replace_none_with_empty_str(row_dict)

    with st.expander("Στοιχεία μαθήματος (πλήρη)"):
        st.write(row_dict)

    
    doc.render(row_dict)
    # doc.save('gen/perigramma.docx')

    # with open('gen/perigramma.docx', "rb") as file:
    #     btn = st.download_button(
    #         label="Download file",
    #         data=file,
    #         file_name=f"perigramma_{docx_code}.docx",
    #         mime="document/docx"
    #     )



# from io import BytesIO
# from docxtpl import DocxTemplate

# # Load the template file
# template = DocxTemplate('my_template.docx')

# # Render the template
# context = {'name': 'John Smith'}
# document = template.render(context)

# # Save the document to a BytesIO buffer
# buffer = BytesIO()
# document.save(buffer)

# # Get the binary data from the buffer
# binary_data = buffer.getvalue()