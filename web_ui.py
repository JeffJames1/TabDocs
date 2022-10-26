"""
    Web UI to parse tableau workbooks and data sources for documentation
"""
import zipfile
import xml.etree.ElementTree as ET
from io import StringIO
import streamlit as st

# import WorkbookDocumentation


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


st.write("#Welcome to the Tableau Workbook Documentation tool!")

infile = st.file_uploader("Tableau workbook to document:")

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

    st.write(root.tag)

    # st.download_button("Download binary file", target_file)
