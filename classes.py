from customtkinter import *
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showerror, showinfo
from tkinter.filedialog import askopenfilenames, askdirectory
from PIL import Image
import comtypes.client
import webbrowser


def string_error(err):
    text = ''
    for txt in err:
        text += f"{txt}"
    return text


class WidgetSet:
    def __init__(self, root: list) -> None:
        self.directory = None
        self.listbox_docx = None
        self.listbox_pdf = None
        self.root = root
        self.doc_to_convert = {}
        self.index = 0
        self.index_path = 0
        self.lists = None

        btn_add_docx = CTkButton(self.root[0], text="Add DOC, DOCX files", width=200, command=self.get_add_files)
        btn_add_docx.place_configure(x=110, y=20)

        btn_add_folder = CTkButton(self.root[0], text="Add convert folder", width=200, command=self.add_folder)
        btn_add_folder.place_configure(x=645, y=20)

        img = CTkImage(light_image=Image.open("convert.ico"), size=(48, 48))
        btn_convert = CTkButton(self.root[0], image=img, text="", width=48, height=48, command=self.convert)
        btn_convert.place_configure(x=463, y=320)

        self.menu = Menu(tearoff=0)
        self.menu.add_command(label="open file", command=self.open_file)
        self.menu.add_command(label="rename in list")
        self.menu.add_command(label="delete from list")

        self.path_label = CTkLabel(self.root[0], text="", width=350, height=40, text_color="black",
                                   fg_color="bisque2", font=("Calibri", 10), corner_radius=10)
        self.path_label.place_configure(x=550, y=690)

        self.docx_label = CTkLabel(self.root[0], text="0 files selected", width=350, height=40, text_color="black",
                                   fg_color="bisque2", font=("Calibri", 10), corner_radius=10)
        self.docx_label.place_configure(x=20, y=690)

        self.convert_text = CTkLabel(self.root[0], text="Convert", width=62, height=7, text_color="black",
                                     fg_color="white", font=("Calibri", 10), corner_radius=10)
        self.convert_text.place_configure(x=463, y=400)

    def get_add_files(self):
        error = []
        files = askopenfilenames()
        for file in files:
            if file[file.rindex("."):] not in ['.docx', '.doc']:
                error.append(f"{file[file.rindex('/') + 1:]}\n")
            else:
                self.doc_to_convert[file] = file[file.rindex("/") + 1:]
        if error:
            showerror(title="WARRNING!", message=f"This files dont convert:\nRename to .docx\n{string_error(error)}")

        if self.listbox_docx is None and self.doc_to_convert != {}:
            self.listbox_docx = Listbox(self.root[1], width=51, height=30, selectmode=BROWSE, font=("Calibri", 12),
                                        bg="aquamarine", fg="black")
            self.listbox_docx.pack(side=LEFT, fill=Y)
            scrollbar = ttk.Scrollbar(self.root[1], orient="vertical", command=self.listbox_docx.yview)
            scrollbar.pack(side=RIGHT, fill=Y)
            self.listbox_docx["yscrollcommand"] = scrollbar.set
            self.listbox_docx.bind('<Button-3>', self.popup_docx)
        elif self.listbox_docx is None and self.doc_to_convert == {}:
            return

        if self.listbox_docx is not None and self.listbox_docx.size() != 0:
            self.listbox_docx.delete(0, END)
            if self.listbox_pdf is not None and self.listbox_pdf.size() != 0:
                self.listbox_pdf.delete(0, END)

        for key, value in self.doc_to_convert.items():
            self.listbox_docx.insert(END, value)
            self.index += 1

        self.docx_label.configure(text=f"{len(self.doc_to_convert)} files selected")
        self.listbox_docx.select_set(0, 0)
        self.listbox_docx.focus_set()

    def popup_docx(self, event):
        self.menu.post(event.x_root, event.y_root)
        self.lists = ['docx', self.listbox_docx]

    def popup_pdf(self, event):
        self.menu.post(event.x_root, event.y_root)
        self.lists = ['pdf', self.listbox_pdf]

    def add_folder(self):
        self.directory = askdirectory()
        self.path_label.configure(text=self.directory)
        self.path_label.update()

    def convert(self):
        error = []
        if self.directory is None:
            showerror(title="WARRNING", message="Add directory for convert")
            return
        if self.doc_to_convert == {}:
            showerror(title="WARRNING", message="Add DOCX files for convert")
            return
        if self.listbox_pdf is None:
            self.listbox_pdf = Listbox(self.root[2], width=51, height=30, selectmode=BROWSE, font=("Calibri", 12),
                                       bg="aquamarine", fg="black")
            self.listbox_pdf.pack(side=LEFT, fill=Y)
            scrollbar = ttk.Scrollbar(self.root[2], orient="vertical", command=self.listbox_pdf.yview)
            scrollbar.pack(side=RIGHT, fill=Y)
            self.listbox_pdf["yscrollcommand"] = scrollbar.set
            self.listbox_pdf.bind('<Button-3>', self.popup_pdf)
        else:
            self.listbox_pdf.delete(0, END)
        self.convert_text.configure(text="Wait...")
        self.convert_text.update()

        for path, file in self.doc_to_convert.items():
            try:
                file_path = f"{self.directory}/{file[:file.rindex('.')]}.pdf"
                path = path.replace("/", chr(92))
                file_path = file_path.replace("/", chr(92))
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(path)
                doc.SaveAs(file_path, FileFormat=17)
                doc.Close()
                word.Quit()
                self.listbox_pdf.insert(END, f"{file[:file.rindex('.')]}.pdf")
                self.listbox_pdf.itemconfigure(END, fg="black")
                self.listbox_pdf.update()
                self.index_path += 1
            except Exception:
                self.listbox_pdf.insert(END, f"{file}")
                self.listbox_pdf.itemconfigure(END, fg="red")
                self.listbox_pdf.update()
                error.append(f"{file}\n")

        if error:
            showerror(title="WARRNING", message=f"{string_error(error)}dont converted\n"
                                                f"the convertation failed with an error(((")
        else:
            showinfo(title="CONGRATULATIONS", message="All files succesfuly converted!")
        self.index_path, self.index = 0, 0
        self.convert_text.configure(text="Convert")

    def open_file(self):
        ids = self.lists[1].curselection()[0]
        file = self.lists[1].get(ids, ids)[0]
        if self.lists[0] == 'docx':
            for key, word in self.doc_to_convert.items():
                if word == file:
                    file = key
        elif self.lists[0] == 'pdf':
            file = f"{self.directory}/{file}"
        self.lists[1].itemconfigure(ids, fg="green")
        self.lists[1].update()
        webbrowser.open(file, new=0, autoraise=True)
