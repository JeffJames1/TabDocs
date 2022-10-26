"""generic functions used in multiple programs"""
import os
import logging
from tkinter import messagebox
from tkinter import filedialog
from tkinter import StringVar
import configparser


def validate_infile(inEntryTxt, file_or_dir):
    """Generic function to validate input file or directory"""
    valid_input = None

    if isinstance(inEntryTxt, StringVar):
        input_file_dir = inEntryTxt.get()
    else:
        input_file_dir = inEntryTxt

    if file_or_dir.get() == "Directory":
        if os.path.isdir(input_file_dir):
            if not input_file_dir.endswith(os.sep):
                input_file_dir = input_file_dir + os.sep
            good_file_notice("Input Directory", input_file_dir)
            valid_input = True
        else:
            bad_file_notice("Input directory", input_file_dir)
    elif file_or_dir.get() == "File":
        if os.path.isfile(input_file_dir):
            good_file_notice("Input File", input_file_dir)
            valid_input = True
        else:
            bad_file_notice("Input directory", input_file_dir)
    else:
        bad_file_notice("Input directory", input_file_dir)

    if valid_input:
        return input_file_dir
    else:
        return None


def validate_file_or_dir(string_var, caption, file_or_dir):
    """Generic function to validate an arbitrary file or directory"""
    if isinstance(string_var, StringVar):
        name_string = string_var.get()
    else:
        name_string = string_var

    if file_or_dir != "directory" and file_or_dir != "file":
        raise ValueError(
            f'must specify either "file" or "directory". {file_or_dir} is not a valid value'
        )

    if (os.path.isfile(name_string) and file_or_dir == "file") or (
        os.path.isdir(name_string) and file_or_dir == "directory"
    ):
        good_file_notice(caption, name_string)
        return name_string
    else:
        bad_file_notice(caption, name_string)
        return None


def good_file_notice(file_type, filename):
    """Output message to indicate a valid file"""
    notice_string = f"{file_type} is {filename}"
    # print(notice_string)
    logging.info(notice_string)


def bad_file_notice(file_type, filename):
    """Output message to indicate an invalid file"""
    notice_string = f"{file_type} {filename} is invalid"
    messagebox.showinfo("Error", notice_string)
    # print(notice_string)
    logging.warning(notice_string)
    # sys.exit(2)


class ConfigEntry(object):
    """Class to define a config entry"""

    # Change to dataclass?
    def __init__(self, name, text_var):
        self.name = name
        self.text_var = text_var


def open_config(section, config_list, configfile=None):
    """Open configuration file to extract parameters"""
    config = configparser.ConfigParser()
    opts = dict()
    if os.name != "posix":
        opts["filetypes"] = [("configuration files", ".cfg"), ("all files", ".*")]
    if not os.path.isfile(configfile):
        configfile = filedialog.askopenfilename(**opts)

    if configfile:
        config.read(configfile)
        for item in config_list:
            item.text_var.set(config.get(section, item.name))
    else:
        messagebox.showinfo("Cancel", "Open dialog canceled")


def save_config(section, config_list):
    """Save parameters to configuration file"""
    config = configparser.ConfigParser()
    config.add_section(section)
    for item in config_list:
        config.set(section, item.name, item.text_var.get())

    opts = dict()
    print(os.name)
    if os.name != "posix":
        opts["filetypes"] = [("configuration files", ".cfg"), ("all files", ".*")]
        opts["defaultextension"] = ".cfg"

    # changed mode from wb to w during Python 3 uplift
    configfile = filedialog.asksaveasfile(mode="w", **opts)
    if configfile:
        config.write(configfile)
        configfile.close()
    else:
        messagebox.showinfo("Cancel", "Save dialog canceled")
