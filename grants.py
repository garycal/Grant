import os
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import scrolledtext, filedialog
import re
import PyPDF2
import docx
from odf.opendocument import load as odt_load
from odf.text import P
import chardet

# List of words to search for
keywords = [
    "racial inequality", "racial justice", "racially", "racism", "sense of belonging",
    "sexual preferences", "social justice", "socio cultural", "socio economic",
    "sociocultural", "socioeconomic", "status", "stereotypes", "systemic", "trauma",
    "under appreciated", "under represented", "under served", "underrepresentation",
    "underrepresented", "underserved", "undervalued", "victim", "women", 
    "women and underrepresented", "advocate", "advocates", "antiracist", "barrier",
    "barriers", "biased", "biases", "bipoc", "community diversity", "disabilities",
    "disability", "discrimination", "discriminatory", "diverse backgrounds",
    "diverse communities", "diverse community", "diverse group", "diverse groups",
    "diversified", "diversify", "diversifying", "diversity and inclusion",
    "diversity equity", "enhance the diversity", "enhancing diversity",
    "equal opportunity", "equality", "equitable", "equity", "ethnicity", "excluded",
    "female", "females", "fostering inclusivity", "gender", "gender diversity",
    "genders", "hate speech", "hispanic minority", "historically", "implicit bias",
    "implicit biases", "inclusion", "inclusive", "inclusiveness", "inclusivity",
    "increase diversity", "increase the diversity", "indigenous community",
    "inequalities", "inequality", "inequitable", "inequities", "institutional",
    "lgbt", "marginalize", "marginalized", "minorities", "minority", "multicultural",
    "polarization", "political", "prejudice", "privileges", "promoting diversity",
    "race and ethnicity", "racial", "racial diversity"
]

# File processing functions
def extract_text_from_pdf(filepath):
    text = ""
    with open(filepath, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text() + " "
    return text

def extract_text_from_docx(filepath):
    doc = docx.Document(filepath)
    return " ".join([para.text for para in doc.paragraphs])

def extract_text_from_odt(filepath):
    doc = odt_load(filepath)
    return " ".join([text_node.data for text_node in doc.getElementsByType(P)])

def extract_text_from_txt(filepath):
    with open(filepath, "rb") as file:
        raw_data = file.read()
        encoding = chardet.detect(raw_data)["encoding"]
        return raw_data.decode(encoding or "utf-8", errors="ignore")

def extract_text_from_rtf(filepath):
    try:
        import striprtf
        with open(filepath, "r", encoding="utf-8") as file:
            return striprtf.striprtf(file.read())
    except ImportError:
        return "Install `striprtf` package to read RTF files."

def extract_text(filepath):
    ext = os.path.splitext(filepath)[-1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(filepath)
    elif ext == ".docx":
        return extract_text_from_docx(filepath)
    elif ext == ".odt":
        return extract_text_from_odt(filepath)
    elif ext == ".txt":
        return extract_text_from_txt(filepath)
    elif ext == ".rtf":
        return extract_text_from_rtf(filepath)
    else:
        return ""

# Function to search for keywords and display results
def search_keywords(text):
    results = []
    for keyword in keywords:
        for match in re.finditer(rf"\b{re.escape(keyword)}\b", text, re.IGNORECASE):
            start = max(match.start() - 50, 0)
            end = min(match.end() + 50, len(text))
            context = text[start:end].replace("\n", " ")
            results.append(f"{keyword}: \"{context}\"")
    return results

def process_file(filepath):
    text = extract_text(filepath)
    matches = search_keywords(text)
    result_text.config(state=tk.NORMAL)
    result_text.delete(1.0, tk.END)
    if matches:
        result_text.insert(tk.END, "\n".join(matches))
    else:
        result_text.insert(tk.END, "No matches found.")
    result_text.config(state=tk.DISABLED)

# Drag-and-drop function
def on_drop(event):
    filepath = event.data.strip("{}")
    process_file(filepath)

# GUI Setup
root = TkinterDnD.Tk()
root.title("Keyword Search in Documents")

# Instructions
label = tk.Label(root, text="Drag and drop a file here (PDF, DOCX, ODT, TXT, RTF):")
label.pack(pady=5)

# Drop Area
drop_frame = tk.Frame(root, width=400, height=100, bg="lightgray", relief="sunken", bd=2)
drop_frame.pack(pady=10)
drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind("<<Drop>>", on_drop)

# Button to select file
def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Supported files", "*.pdf;*.docx;*.odt;*.txt;*.rtf"),
        ("PDF Files", "*.pdf"),
        ("Word Documents", "*.docx"),
        ("OpenDocument Text", "*.odt"),
        ("Text Files", "*.txt"),
        ("Rich Text Format", "*.rtf")
    ])
    if file_path:
        process_file(file_path)

browse_button = tk.Button(root, text="Select File", command=open_file_dialog)
browse_button.pack(pady=5)

# Scrollable Text Area
result_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
result_text.pack(padx=10, pady=10)
result_text.config(state=tk.DISABLED)

# Run Application
root.geometry("600x500")
root.mainloop()

