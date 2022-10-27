"""
    Web UI to parse tableau workbooks and data sources for documentation
"""
import zipfile
from os.path import splitext, basename
import xml.etree.ElementTree as ET
from io import StringIO, BytesIO
import streamlit as st

import WorkbookDocumentation


def find_file_in_zip(zip_file):
    """Find workbook or data source in zip file"""
    candidate_files = filter(
        lambda x: x.split(".")[-1] in ("twb", "tds"), zip_file.namelist()
    )

    for filename in candidate_files:
        with zip_file.open(filename) as xml_candidate:
            try:
                ET.parse(xml_candidate)
                return filename
            except ET.ParseError:
                # That's not an XML file by gosh
                pass


st.title("Welcome to the Tableau Documentation tool!")
st.write("Update a workbook or data source to be documented")
st.write("twb/twbx and tds/tdsx files are supported")

infile = st.file_uploader("Tableau file to document:")

st.write("Click the 'x' next to the upload file to remove it or upload a new file to restart.")

if infile is not None:
    if zipfile.is_zipfile(infile):
        with zipfile.ZipFile(infile) as zip_object:
            target_file = find_file_in_zip(zip_object)
            with zip_object.open(target_file) as xml_source:
                object_tree = ET.parse(xml_source)
                root = object_tree.getroot()

    else:
        stringio = StringIO(infile.getvalue().decode("utf-8"))
        root = ET.fromstring(stringio.read())

    style_guide = None

    documentation = WorkbookDocumentation.WorkbookDocumentation(root, style_guide)

    doc_workbook = documentation.build_excel_workbook()

    with BytesIO() as output:
        doc_workbook.save(output)
        output.seek(0)
        byte_data = output.read()

    out_file_name = splitext(basename(infile.name))[0] + " Documentation.xlsx"

    st.download_button(
        "Download documentation file", byte_data, file_name=out_file_name
    )
