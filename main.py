import os
import tkinter as tk

import comtypes.client
from tkinterdnd2 import TkinterDnD, DND_FILES


def convert_ppt_to_pdf(input_file_path):
    output_file_path = input_file_path.replace('.ppt', '.pdf').replace('.pptx', '.pdf')

    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    slides = powerpoint.Presentations.Open(input_file_path)
    slides.SaveAs(output_file_path, 32)
    slides.Close()

    powerpoint.Quit()
    print(f'Converted {input_file_path} to {output_file_path}')


def convert_word_to_pdf(input_file_path):
    output_file_path = input_file_path.replace('.doc', '.pdf').replace('.docx', '.pdf')

    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.abspath(output_file_path)

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0

    doc = word.Documents.Open(input_file_path)
    doc.SaveAs(output_file_path, FileFormat=17)  # 17 is the code for wdFormatPDF
    doc.Close()

    word.Quit()
    print(f'Converted {input_file_path} to {output_file_path}')


def handle_drop(event):
    file_path = event.data.strip('{}')
    if file_path.endswith(('.ppt', '.pptx')):
        convert_ppt_to_pdf(file_path)
    elif file_path.endswith(('.doc', '.docx')):
        convert_word_to_pdf(file_path)
    else:
        print(f'Unsupported file type: {file_path}')


def main():
    root = TkinterDnD.Tk()
    root.title('Drag and Drop PDF Converter')
    root.geometry('400x200')

    label = tk.Label(root, text='Drag and drop your PPT, PPTX, DOC, or DOCX files here')
    label.pack(pady=20)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', handle_drop)

    root.mainloop()


if __name__ == "__main__":
    main()
