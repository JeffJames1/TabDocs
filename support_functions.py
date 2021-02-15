import os
import logging
from tkinter import messagebox
import configparser
from tkinter import filedialog
try:
    from Tkinter import *
except ImportError:
    from tkinter import *


def validate_infile(inEntryTxt, file_or_dir):
    """Generic function to validate input file or directory"""
    valid_input = None

    if isinstance(inEntryTxt, StringVar):
        input_file_dir = inEntryTxt.get()
    else:
        input_file_dir = inEntryTxt

    if file_or_dir.get() == 'Directory':
        if os.path.isdir(input_file_dir):
            if not input_file_dir.endswith(os.sep):
                inputfiledir = input_file_dir + os.sep
            good_file_notice("Input Directory", input_file_dir)
            valid_input = True
        else:
            bad_file_notice("Input directory", input_file_dir)
    elif file_or_dir.get() == 'File':
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
        raise ValueError('must specify either "file" or "directory". {} is not a valid value'.format(file_or_dir))

    if (os.path.isfile(name_string) and file_or_dir == "file") or \
            (os.path.isdir(name_string) and file_or_dir == "directory"):
        good_file_notice(caption, name_string)
        return name_string
    else:
        bad_file_notice(caption, name_string)
        return None


def good_file_notice(file_type, filename):
    """Output message to indicate a valid file"""
    notice_string = '{} is {}'.format(file_type, filename)
    # print(notice_string)
    logging.info(notice_string)


def bad_file_notice(file_type, filename):
    """Output message to indicate an invalid file"""
    notice_string = '{} {} is invalid'.format(file_type, filename)
    messagebox.showinfo("Error", notice_string)
    # print(notice_string)
    logging.warning(notice_string)
    # sys.exit(2)


class config_entry(object):
    def __init__(self, name, text_var):
        self.name = name
        self.text_var = text_var


def open_config(section, config_list, configfile=None):

    config = configparser.ConfigParser()
    opts = dict()
    if os.name != 'posix':
        opts['filetypes'] = [('configuration files', '.cfg'), ('all files', '.*')]
    if not os.path.isfile(configfile):
        configfile = filedialog.askopenfilename(**opts)

    if configfile:
        config.read(configfile)
        for item in config_list:
            item.text_var.set(config.get(section, item.name))
    else:
        messagebox.showinfo("Cancel", "Open dialog canceled")


def save_config(section, config_list):

    config = configparser.ConfigParser()
    config.add_section(section)
    for item in config_list:
        config.set(section, item.name, item.text_var.get())

    opts = dict()
    print(os.name)
    if os.name != 'posix':
        opts['filetypes'] = [('configuration files', '.cfg'), ('all files', '.*')]
        opts['defaultextension'] = ".cfg"

    # changed mode from wb to w during Python 3 uplift
    configfile = filedialog.asksaveasfile(mode='w', **opts)
    if configfile:
        config.write(configfile)
        configfile.close()
    else:
        messagebox.showinfo("Cancel", "Save dialog canceled")
