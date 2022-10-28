"""
    Web UI to parse tableau workbooks and data sources for documentation
"""
import zipfile
from os.path import splitext, basename
import xml.etree.ElementTree as ET
from io import StringIO, BytesIO
import streamlit as st
from WorkbookDocumentation import WorkbookDocumentation


def find_file_in_zip(zip_file) -> str:
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


def generate_xml_root(infile) -> ET.Element:
    """Get the root element of the object XML"""
    if zipfile.is_zipfile(infile):
        with zipfile.ZipFile(infile) as zip_object:
            target_file = find_file_in_zip(zip_object)
            with zip_object.open(target_file) as xml_source:
                object_tree = ET.parse(xml_source)
                root = object_tree.getroot()

    else:
        stringio = StringIO(infile.getvalue().decode("utf-8"))
        root = ET.fromstring(stringio.read())
    return root


def convert_to_bytes(doc_workbook) -> BytesIO:
    """Convert object to byte data"""
    with BytesIO() as output:
        doc_workbook.save(output)
        output.seek(0)
        byte_data = output.read()
    return byte_data


def process_file(uploaded_file) -> BytesIO:
    """process file to generate documentation workbook"""
    root = generate_xml_root(uploaded_file)
    style_guide = None
    documentation = WorkbookDocumentation(root, style_guide)
    doc_workbook = documentation.build_excel_workbook()
    byte_data = convert_to_bytes(doc_workbook)
    return byte_data


def process_uploaded_files(uploaded_files):
    """process list of files and generate document files"""
    file_archive = BytesIO()
    with zipfile.ZipFile(
        file_archive, "a", compression=zipfile.ZIP_DEFLATED, allowZip64=False
    ) as zip_file:
        for uploaded_file in uploaded_files:
            byte_data = process_file(uploaded_file)
            out_file_name = (
                splitext(basename(uploaded_file.name))[0] + " Documentation.xlsx"
            )
            zip_file.writestr(out_file_name, byte_data)
    return file_archive, byte_data, out_file_name


def main():
    """Render page to create documentation"""
    st.title("Welcome to the Tableau Documentation tool!")
    st.write("Update one or more workbooks and/or data sources to be documented")

    with st.form("documentation-form", clear_on_submit=True):
        uploaded_files = st.file_uploader(
            "Tableau file(s) to document:",
            type=["twb", "twbx", "tds", "tdsx"],
            accept_multiple_files=True,
        )

        submitted = st.form_submit_button("Process file(s)")

    if submitted is not None and len(uploaded_files) != 0:
        file_archive, byte_data, out_file_name = process_uploaded_files(uploaded_files)
        if len(uploaded_files) == 1:
            st.download_button(
                "Download documentation file", byte_data, file_name=out_file_name
            )
        elif len(uploaded_files) > 1:
            st.download_button("Download zip file", file_archive, "Documentation.zip")


if __name__ == "__main__":
    main()
