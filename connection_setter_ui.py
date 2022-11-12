"""
    Web UI to update connection information in tableau workbooks and data sources
"""
import zipfile
from os.path import splitext, basename
import xml.etree.ElementTree as ET
from io import StringIO, BytesIO
import re
from xml.dom import minidom
import streamlit as st


def set_server_data_source_connection(root, server_info):
    """Apply server data source information"""
    # st.write("set_server_data_source_connection")
    for connection in root.findall(".//datasource/connection[@class='sqlproxy']"):
        connection.set("server", server_info["tableau_server"])
        connection.set("site", server_info["tableau_site"])

    return root


def detect_custom_sql(root):
    """look for relation nodes with text that starts with 'select'.
    Warn user that manual work may be needed to fix
    """
    # st.write("detect_custom_sql")
    xpath_string = (
        ".//relation[starts-with(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        + "'abcdefghijklmnopqrstuvwxyz'), 'select ')]/text()"
    )
    for custom_sql in root.xpath(xpath_string):
        st.warning("Potential custom SQL - manual correction may be needed", icon="⚠️")
        st.warning(custom_sql)


def set_database_schema(root, server_info):
    """Apply database schema information"""
    # st.write("set_database_schema")
    for relation in root.findall(".//relation[@table]"):
        table_name = relation.get("table")
        if re.search(r"\[(.*)\]\.\[(.*)\]", table_name):
            parsed_name_list = re.search(r"\[(.*)\]\.\[(.*)\]", table_name)
            table = parsed_name_list.group(2)
            schema_table = f"[{server_info['database_schema']}].[{table}]"
            relation.set("table", schema_table)
    for relation in root.findall(
        ".//_.fcp.ObjectModelEncapsulateLegacy.false...relation[@table]"
    ):
        table_name = relation.get("table")
        if re.search(r"\[(.*)\]\.\[(.*)\]", table_name):
            parsed_name_list = re.search(r"\[(.*)\]\.\[(.*)\]", table_name)
            table = parsed_name_list.group(2)
            schema_table = f"[{server_info['database_schema']}].[{table}]"
            relation.set("table", schema_table)
    for relation in root.findall(
        ".//_.fcp.ObjectModelEncapsulateLegacy.true...relation[@table]"
    ):
        table_name = relation.get("table")
        if re.search(r"\[(.*)\]\.\[(.*)\]", table_name):
            parsed_name_list = re.search(r"\[(.*)\]\.\[(.*)\]", table_name)
            table = parsed_name_list.group(2)
            schema_table = f"[{server_info['database_schema']}].[{table}]"
            relation.set("table", schema_table)

    return root


def set_database_connection(root, server_info):
    """Apply database connection information"""
    # st.write("set_database_connection")
    for connection in root.findall(
        ".//named-connections/named-connection/connection[@class='vertica']"
    ):
        connection.set("server", server_info["database_server"])
        connection.set("schema", server_info["database_schema"])
        connection.set("username", server_info["database_user"])

    return root


def find_file_in_zip(zip_file) -> str:
    """Find workbook or data source in zip (i.e. packaged) file"""
    # st.write("find_file_in_zip")
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
    # st.write("generate_xml_root")
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


def convert_to_bytes(root: ET.Element) -> BytesIO:
    """Convert ElementTree element to byte data"""
    # st.write("convert_to_bytes")
    with BytesIO() as output:
        xmlstr = minidom.parseString(ET.tostring(root, "utf-8")).toprettyxml(
            indent="    "
        )
        output.write(bytes(xmlstr, "utf-8"))
        output.seek(0)
        byte_data = output.read()
        # st.write(xmlstr)
    return byte_data


def process_file(uploaded_file, server_info: dict) -> BytesIO:
    """process file to apply connection information"""
    # st.write("process_file")
    root = generate_xml_root(uploaded_file)
    if (
        server_info.get("database_schema")
        and server_info.get("database_server")
        and server_info.get("database_user")
    ):
        root = set_database_connection(root, server_info)
        root = set_database_schema(root, server_info)
        # detect_custom_sql(root)
    root = set_server_data_source_connection(root, server_info)
    # st.write(type(root))

    # todo: still need to package XML and return it
    # return tree as byte data if not a zipfile
    # write file into zip if packaged
    if not zipfile.is_zipfile(uploaded_file):
        byte_data = convert_to_bytes(root)
    else:
        st.write("Not handling packaged files yet. Sorry!")
    return byte_data


def process_uploaded_files(uploaded_files: list, server_info: dict):
    """process list of files and generate document files"""
    file_archive = BytesIO()
    with zipfile.ZipFile(
        file_archive, "a", compression=zipfile.ZIP_DEFLATED, allowZip64=False
    ) as zip_file:
        for uploaded_file in uploaded_files:
            byte_data = process_file(uploaded_file, server_info)
            out_file_name = (
                splitext(basename(uploaded_file.name))[0]
                + " updated."
                + splitext(basename(uploaded_file.name))[1]
            )
            zip_file.writestr(out_file_name, byte_data)
    return file_archive, byte_data, out_file_name


def main():
    """Render page to set connections"""

    server_info = {}

    st.title("Welcome to the Tableau Connection Setter!")
    st.write("Update one or more workbooks and/or data sources to be configured")

    with st.form("connection-form", clear_on_submit=False):
        uploaded_files = st.file_uploader(
            "Tableau file(s) to configure:",
            type=["twb", "twbx", "tds", "tdsx"],
            accept_multiple_files=True,
        )
        st.write("Tableau Server data")
        server_info["tableau_server"] = st.text_input("Tableau Server")
        server_info["tableau_site"] = st.text_input("Site ID")
        st.write("Database information for direct connections")
        server_info["database_server"] = st.text_input("Database Server")
        server_info["database_schema"] = st.text_input("Database Schema")
        server_info["database_user"] = st.text_input("Database User")

        submitted = st.form_submit_button("Process file(s)")

    if submitted is not None and len(uploaded_files) != 0:
        file_archive, byte_data, out_file_name = process_uploaded_files(
            uploaded_files, server_info
        )
        if len(uploaded_files) == 1:
            st.download_button(
                "Download updated file", byte_data, file_name=out_file_name
            )
        elif len(uploaded_files) > 1:
            st.download_button("Download zip file", file_archive, "updated files.zip")
        # st.write("Success!!!")
    else:
        st.warning("Must select a file or files for processing")


if __name__ == "__main__":
    main()
