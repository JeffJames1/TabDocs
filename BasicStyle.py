from tkinter import *
from tkinter.ttk import *


class BasicStyle:

    def create_row(self, frame, row, label_text, entry_variable, button_function=None, button_label=None, password='no'):
        label = self.create_label(frame, row, 1, label_text)
        entry = self.create_entry(frame, row, 2, entry_variable, password)
        if button_function is not None and button_label is not None:
            button = self.create_button(frame, row, 3, button_function, button_label, 12)
            return label, entry, button
        else:
            return label, entry

    @staticmethod
    def create_menu(top, open, save, quit):
        menubar = Menu(top, font="TkMenuFont")
        top.configure(menu=menubar)
        file = Menu(top, tearoff=0)
        menubar.add_cascade(
            menu=file,
            activebackground="#d9d9d9",
            activeforeground="#000000",
            background="#d9d9d9",
            font="TkMenuFont",
            foreground="#000000",
            label="File")
        file.add_command(
            activebackground="#d8d8d8",
            activeforeground="#000000",
            background="#d9d9d9",
            command=open,
            font="TkMenuFont",
            foreground="#000000",
            label="Open")
        file.add_command(
            activebackground="#d8d8d8",
            activeforeground="#000000",
            background="#d9d9d9",
            command=save,
            font="TkMenuFont",
            foreground="#000000",
            label="Save")
        file.add_separator(
            background="#d9d9d9")
        file.add_command(
            activebackground="#d8d8d8",
            activeforeground="#000000",
            background="#d9d9d9",
            command=quit,
            font="TkMenuFont",
            foreground="#000000",
            label="Quit")
        return file

    @staticmethod
    def create_frame(top):
        frame = Frame(top)
        frame.grid(column=0, row=0, sticky=(N, W, E, S))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.pack(pady=20, padx=20)
        return frame

    @staticmethod
    def create_rdo(frame, column, text, value, variable):
        rdo = Radiobutton(frame)
        rdo.grid(row=0, column=column)
        rdo.configure(text=text)
        rdo.configure(value=value)
        rdo.configure(variable=variable)
        return rdo

    @staticmethod
    def create_button(frame, row, column, command, text, size=0, sticky=None):
        button = Button(frame)
        button.grid(row=row, column=column, sticky=sticky)
        button.configure(command=command)
        button.configure(text=text)
        button.update_idletasks()
        return button

    @staticmethod
    def create_label(frame, row, column, text):
        label = Label(frame)
        label.grid(row=row, column=column, sticky=W)
        label.configure(text=text)
        label.configure(padding=5)
        return label

    @staticmethod
    def create_entry(frame, row, column, variable, password='no'):
        if password == 'yes':
            show = '*'
        else:
            show = ''
        entry = Entry(frame, show=show)
        entry.grid(row=row, column=column)
        entry.configure(textvariable=variable)
        return entry

    def create_file_dir_frame(self, frame, row, fileordir):
        labelframe = LabelFrame(frame)
        labelframe.grid(row=row, column=2, rowspan=2)
        labelframe.configure(relief=GROOVE)
        labelframe.configure(text='''Scope to Process''')
        labelframe.configure(width=170)
        labelframe.configure(padding=5)
        rdoFile = self.create_rdo(labelframe, 0, '''File''', "File", fileordir)
        rdoFile = self.create_rdo(labelframe, 1, '''Directory''', "Directory", fileordir)
        return labelframe, rdoFile

    def create_ok_quit(self, frame, ok, quit):
        buttonframe = Frame(frame)
        buttonframe.grid(row=100, columnspan=10, rowspan=2)
        buttonframe.configure(padding=5)
        btnOK = self.create_button(buttonframe, 1, 1, ok, '''OK''', 12)
        btnQuit = self.create_button(buttonframe, 1, 2, quit, '''Quit''', 12)