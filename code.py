import tkinter as tk
from tkinter import filedialog, scrolledtext
import pdfplumber
import docx
import nltk
import pyttsx3
import threading
import pytesseract
from PIL import Image

# Configure Tesseract-OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Ensure NLTK tokenizer is available
nltk.download("punkt")

# Initialize text-to-speech engine
engine = pyttsx3.init()
engine.setProperty("rate", 150)  # Adjust speech speed

# Global variables
extracted_text = ""
summarized_text = ""
is_speaking = False
stop_requested = False

# Initialize GUI
root = tk.Tk()
root.title("AI Text Summarizer with Speech")
root.geometry("900x650")

# Function to extract text from files
def extract_text_from_file():
    global extracted_text
    file_path = filedialog.askopenfilename(filetypes=[
        ("All Supported Files", "*.jpg;*.png;*.pdf;*.docx;*.txt"),
        ("Image Files", "*.jpg;*.png"),
        ("PDF Files", "*.pdf"),
        ("Word Files", "*.docx"),
        ("Text Files", "*.txt")
    ])
    
    if not file_path:
        return

    extracted_text_display.delete("1.0", tk.END)
    summarized_text_display.delete("1.0", tk.END)

    try:
        if file_path.lower().endswith((".jpg", ".png")):
            extracted_text = extract_text_from_image(file_path)
        elif file_path.lower().endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
        elif file_path.lower().endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
        elif file_path.lower().endswith(".txt"):
            extracted_text = extract_text_from_txt(file_path)
        else:
            extracted_text = "Unsupported file format!"

        extracted_text_display.insert(tk.END, extracted_text)
        summarize_text()
    except Exception as e:
        extracted_text_display.insert(tk.END, f"Error: {str(e)}")

# Extraction functions
def extract_text_from_image(file_path):
    image = Image.open(file_path)
    return pytesseract.image_to_string(image).strip()

def extract_text_from_pdf(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text.strip() if text else "No text found in PDF!"

def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs]).strip() or "No text found in DOCX!"

def extract_text_from_txt(file_path):
    with open(file_path, "r", encoding="utf-8") as file:
        return file.read().strip() or "No text found in TXT!"

# Summarization function
def summarize_text():
    global summarized_text
    summarized_text_display.delete("1.0", tk.END)
    
    sentences = nltk.sent_tokenize(extracted_text)
    num_sentences = max(2, len(sentences) // 3)
    summarized_text = " ".join(sentences[:num_sentences])
    
    summarized_text_display.insert(tk.END, summarized_text)

def resummarize_text():
    """Re-summarizes the extracted text differently."""
    global summarized_text
    summarized_text_display.delete("1.0", tk.END)

    sentences = nltk.sent_tokenize(extracted_text)
    num_sentences = max(3, len(sentences) // 2)  # Different summarization method
    summarized_text = " ".join(sentences[:num_sentences])

    summarized_text_display.insert(tk.END, summarized_text)

def start_speaking():
    global is_speaking, stop_requested, engine
    if not is_speaking:
        is_speaking = True
        stop_requested = False
        engine = pyttsx3.init()  # Reinitialize engine to avoid issues after stopping
        threading.Thread(target=speak_text, daemon=True).start()

def speak_text():
    """Reads the summarized text out loud while highlighting sentences."""
    global is_speaking, stop_requested
    is_speaking = True

    sentences = nltk.sent_tokenize(summarized_text)
    summarized_text_display.tag_remove("highlight", "1.0", tk.END)  # Remove old highlights

    for sentence in sentences:
        if stop_requested:
            break
        
        # Remove previous highlight
        summarized_text_display.tag_remove("highlight", "1.0", tk.END)
        
        # Find the sentence position in the text and highlight it
        start = summarized_text_display.search(sentence, "1.0", tk.END)
        if start:
            end = f"{start}+{len(sentence)}c"
            summarized_text_display.tag_add("highlight", start, end)
            summarized_text_display.tag_config("highlight", background="yellow")
            summarized_text_display.see(start)  # Scroll to the highlighted sentence

        engine.say(sentence)
        engine.runAndWait()

    is_speaking = False
    summarized_text_display.tag_remove("highlight", "1.0", tk.END)  # Remove highlights after speaking

def stop_speaking():
    """Stops the speech immediately and clears highlighting."""
    global stop_requested, is_speaking, engine
    stop_requested = True
    is_speaking = False
    engine.stop()
    engine = None  # Reset engine to prevent issues
    summarized_text_display.tag_remove("highlight", "1.0", tk.END)  # Remove highlights

# UI Elements
frame = tk.Frame(root)
frame.pack(pady=10)

upload_btn = tk.Button(frame, text="Upload File", command=extract_text_from_file, bg="lightblue", fg="black", width=20)
upload_btn.pack()

tk.Label(root, text="Extracted Text:", font=("Arial", 12)).pack()
extracted_text_display = scrolledtext.ScrolledText(root, height=12, width=100, wrap=tk.WORD)
extracted_text_display.pack(padx=10, pady=5)

tk.Label(root, text="Summarized Text:", font=("Arial", 12)).pack()
summarized_text_display = scrolledtext.ScrolledText(root, height=8, width=100, wrap=tk.WORD)
summarized_text_display.pack(padx=10, pady=5)

# Button Frame
btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

speak_btn = tk.Button(btn_frame, text="Speak", command=start_speaking, bg="green", fg="white", width=15)
speak_btn.grid(row=0, column=0, padx=10)

stop_btn = tk.Button(btn_frame, text="Stop", command=stop_speaking, bg="red", fg="white", width=15)
stop_btn.grid(row=0, column=1, padx=10)

resummarize_btn = tk.Button(btn_frame, text="Re-Summarize", command=resummarize_text, bg="orange", fg="black", width=15)
resummarize_btn.grid(row=0, column=2, padx=10)

root.mainloop()
