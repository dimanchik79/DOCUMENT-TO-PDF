
from customtkinter import *
from classes import WidgetSet


def main():
    color = ["CornflowerBlue", "aquamarine"]
    root = CTk()
    root.title("COVERT DOCUMENTS TO PDF")
    root.iconbitmap(default="icon.ico")

    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{800}x{600}+{int(sw / 2 - 400)}+{int(sh / 2 - 300)}")

    root.resizable(False, False)
    root.configure(fg_color=color[0])

    frame_docx = CTkFrame(root, height=490, width=350, border_width=1, fg_color=color[1], border_color="black")
    frame_docx.place_configure(x=20, y=70)

    frame_pdf = CTkFrame(root, height=490, width=350, border_width=1, fg_color=color[1], border_color="black")
    frame_pdf.place_configure(x=550, y=70)

    WidgetSet(root=[root, frame_docx, frame_pdf])

    root.mainloop()


if __name__ == "__main__":
    main()
