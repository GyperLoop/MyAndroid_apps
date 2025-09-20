import os
from PyPDF2 import PdfReader
from docx import Document
from openpyxl import load_workbook

from kivy.app import App
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.lang import Builder

Builder.load_file("document_reader.kv")

def read_pdf(path: str) -> str:
    text = []
    with open(path, 'rb') as f:
        reader = PdfReader(f)
        for page in reader.pages:
            text.append(page.extract_text() or "")
    return "\n".join(text)

def read_docx(path: str) -> str:
    doc = Document(path)
    return "\n".join([para.text for para in doc.paragraphs])

def read_txt(path: str) -> str:
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()
    
def read_xlsx(path: str) -> str:
    wb = load_workbook(path, data_only=True)
    output = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        output.append(f"--- Sheet: {sheet_name} ---")
        for row in sheet.iter_rows(value_only=True):
            row_values = [str(cell) if cell is not None else "" for cell in row] 
            output.append("\t".join(row_values))
        output.append("")
    return "\n".join(output)

def read_document(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == '.pdf':
        return read_pdf(path)
    elif ext == '.docx':
        return read_docx(path)
    elif ext == '.txt':
        return read_txt(path)
    elif ext == '.xlsx':
        return read_xlsx(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    
class DocumentReaderApp(App):
    def build(self):
        return self.root
    
    def open_file_chooser(self):
        layout = BoxLayout(orientation='vertical', spacing=5)
        chooser = FileChooserListView(filters=['*.pdf', '*.docx', '*.txt', '*.xlsx'])
        btn = Button(text='Open Selected', size_hint_y=None, height='48dp')

        def open_selected(instance):
            if chooser.selection:
                file_path = chooser.selection[0]
                try:
                    content = read_document(file_path)
                    self.root.ids.output.text = content
                except Exception as e:
                    self.root.ids.output.text = f"Error: {e}"
                popup.dismiss()

        btn.bind(on_relese=open_selected)
        layout.add_widget(chooser)
        layout.add_widget(btn)
        popup = Popup(title="Select a Document",
                      content=layout,
                      size_hint=(0.9, 0.9))
        popup.open()
    
if __name__ == '__main__':
    DocumentReaderApp().run()