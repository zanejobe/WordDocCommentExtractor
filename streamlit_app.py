"""
modified from this page https://carstenknoch.com/2020/02/qualitative-data-analysis-with-microsoft-word-comments-python-updated/
"""

import streamlit as st

import zipfile
import csv
import lxml
from bs4 import BeautifulSoup as Soup
import tkinter as tk
from tkinter import filedialog
import re

# SETTING PAGE CONFIG TO WIDE MODE
st.set_page_config(layout="wide", page_title="Word Comment Extractor (by Zane Jobe)")

csvw=None

@st.cache
def load_data(uploaded_file):
	if uploaded_file is not None:
		with open('output.csv', 'w', newline='', encoding='utf-8-sig') as f:
			csvw = csv.writer(f)
			unzip = zipfile.ZipFile(uploaded_file)
			comments = Soup(unzip.read('word/comments.xml'), 'html.parser')
			doc = unzip.read('word/document.xml').decode()
			start_loc = {x.group(1): x.start() for x in re.finditer(r'<w:commentRangeStart.*?w:id="(.*?)"', doc)}
			end_loc = {x.group(1): x.end() for x in re.finditer(r'<w:commentRangeEnd.*?w:id="(.*?)".*?>', doc)}
			for c in comments.find_all('w:comment'):
				c_id = c.attrs['w:id']
				xml = re.sub(r'(<w:p .*?>)', r'\1 ', doc[start_loc[c_id]:end_loc[c_id] + 1])
				csvw.writerow([''.join(c.findAll(text=True)), ''.join(Soup(xml, 'lxml').findAll(text=True))])
			unzip.close()
		return csvw

# from here https://discuss.streamlit.io/t/heres-a-download-function-that-works-for-dataframes-and-txt/4052
def download_link(csvw, download_filename, download_link_text):
    """
    Generates a link to download the given object_to_download.

    object_to_download (str, pd.DataFrame):  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.csv, some_txt_output.txt
    download_link_text (str): Text to display for download link.

    Examples:
    download_link(YOUR_DF, 'YOUR_DF.csv', 'Click here to download data!')
    download_link(YOUR_STRING, 'YOUR_STRING.txt', 'Click here to download your text!')

    """
    if isinstance(object_to_download,pd.DataFrame):
        object_to_download = object_to_download.to_csv(index=False)

    # some strings <-> bytes conversions necessary here
    b64 = base64.b64encode(object_to_download.encode()).decode()

    return f'<a href="data:file/txt;base64,{b64}" download="{download_filename}">{download_link_text}</a>'

# DO STUFF
# Load data
st.write('# Word Comment Extractor')
st.write('To begin using the app, load your Word doc (.docx) using the file upload option below.')
uploaded_file = st.sidebar.file_uploader(' ', type=['.docx'])

load_data(uploaded_file)
if csvw is not None:
	st.write(csvw)

if st.button('Download as CSV'):
    tmp_download_link = download_link(df, 'output.csv', 'Click here to download your data!')
    st.markdown(tmp_download_link, unsafe_allow_html=True)
