"""
Extract data from Tableau packaged workbooks. 
XML must be removed from the archive and then reconstructed
"""
import contextlib
import os
import shutil
import tempfile
import zipfile

# import lxml.etree as
import xml.etree.ElementTree as ET
from pathlib import Path


def xml_open(filename):
    """Opens the provided 'filename'. Handles detecting if the file is an archive,
    detecting the document version, and validating the root tag."""

    # Is the file a zip (.twbx or .tdsx)
    if zipfile.is_zipfile(filename):
        tree, contents = get_xml_from_archive(filename)
    else:
        print(Path(filename).name)
        with open(filename, encoding="utf-8") as f:
            contents = f.read()
        try:
            tree = ET.parse(filename)

        except ET.ParseError:
            # Fix to deal with namespace problem in application Data Models
            # can be removed when Data Models are fixed
            with temporary_directory() as temppath:
                fixedfile = fix_namespace(filename, temppath)
                tree = ET.parse(fixedfile)

    return tree, contents


@contextlib.contextmanager
def temporary_directory(*args, **kwargs):
    """Create temporary directory and delete when finished"""
    d = tempfile.mkdtemp(*args, **kwargs)
    try:
        yield d
    finally:
        shutil.rmtree(d)


def find_file_in_zip(zip_file):
    """Returns the twb/tds file from a Tableau packaged file format. Packaged
    files can contain cache entries which are also valid XML, so only look for
    files with a .tds or .twb extension.
    """

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


def get_xml_from_archive(filename):
    """Extract workbook xml from archive"""
    with zipfile.ZipFile(filename) as zf:
        with zf.open(find_file_in_zip(zf), "r") as xml_file:
            with temporary_directory() as temppath:
                # print(os.path.join(temppath, "temp_file"))
                temp_file_name = os.path.join(temppath, "temp_file")
                # print("xml file: {}".format(xml_file))
                with open(temp_file_name, "a", encoding="utf-8") as temp_file:
                    for line in xml_file:
                        temp_file.write(line.decode("utf-8"))
                xml_contents = xml_file.read()

                try:
                    with open(temp_file_name, encoding="utf-8") as file:
                        xml_tree = ET.parse(file)
                except ET.ParseError:
                    # Fix to deal with namespace problem in application Data Models
                    fixedfile = fix_namespace(temp_file_name, temppath)
                    with open(fixedfile, encoding="utf-8") as file:
                        xml_tree = ET.parse(file)

    return xml_tree, xml_contents


def build_archive_file(archive_contents, zip_file):
    """Build a Tableau-compatible archive file."""

    # This is tested against Desktop and Server, and reverse engineered by lots
    # of trial and error. Do not change this logic.
    for root_dir, _, files in os.walk(archive_contents):
        relative_dir = os.path.relpath(root_dir, archive_contents)
        for f in files:
            temp_file_full_path = os.path.join(archive_contents, relative_dir, f)
            zipname = os.path.join(relative_dir, f)
            zip_file.write(temp_file_full_path, arcname=zipname)


def save_into_archive(xml_tree, filename, new_filename=None):
    """
    Saving an archive means extracting the contents into a temp folder,
    saving the changes over the twb/tds in that folder, and then
    packaging it back up into a zip with a very specific format
    e.g. no empty files for directories, which Windows and Mac do by default
    """

    if new_filename is None:
        new_filename = filename

    # Extract to temp directory
    with temporary_directory() as temp_path:
        with zipfile.ZipFile(filename) as zf:
            xml_file = find_file_in_zip(zf)
            zf.extractall(temp_path)
        # Write the new version of the file to the temp directory
        xml_tree.write(
            os.path.join(temp_path, xml_file),
            encoding="utf-8",
            pretty_print=True,
            xml_declaration=True,
        )

        # Write the new archive with the contents of the temp folder
        with zipfile.ZipFile(
            new_filename, "w", compression=zipfile.ZIP_DEFLATED
        ) as new_archive:
            build_archive_file(temp_path, new_archive)


def save_file(container_file, xml_tree, new_filename=None):
    """save xml to file"""

    if new_filename is None:
        new_filename = container_file

    if zipfile.is_zipfile(container_file):
        save_into_archive(xml_tree, container_file, new_filename)
    else:
        xml_tree.write(
            new_filename, encoding="utf-8", pretty_print=True, xml_declaration=True
        )


def fix_namespace(filename, outpath):
    """Namespace issues within the xml. Fix so that parsing works correctly"""
    replacements = {" user:": " "}

    # print("File exists: {}".format(os.path.isfile(filename)))
    # with open(filename,'r') as temp_file:
    #     for line in temp_file:
    #         print(line)
    outfile_name = os.path.join(outpath, Path(filename).name)
    tempfile_name = os.path.join(outpath, "tempfile")

    shutil.copyfile(filename, tempfile_name)

    with open(tempfile_name, "r", encoding="utf-8") as infile, open(
        outfile_name, "w", encoding="utf-8"
    ) as outfile:
        for line in infile:
            for src, target in replacements.items():
                line = line.replace(src, target)
            outfile.write(line)

    os.remove(tempfile_name)

    return outfile_name
