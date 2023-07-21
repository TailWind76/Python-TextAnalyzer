import re
from collections import Counter
import tkinter as tk
from tkinter import filedialog
import docx

def preprocess_text(text):
    text = re.sub(r'[^\w\s]', '', text.lower())
    words = text.split()
    return words

def analyze_text(text):
    words = preprocess_text(text)
    word_count = Counter(words)
    return word_count

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("Word Files", "*.docx")])
    if file_path:
        if file_path.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
        elif file_path.endswith('.docx'):
            doc = docx.Document(file_path)
            content = " ".join(paragraph.text for paragraph in doc.paragraphs)
        
        word_count = analyze_text(content)
        display_results(word_count)

def display_results(word_count):
    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, "Результаты анализа:\n============================\n")
    for word, count in word_count.most_common(10):
        result_text.insert(tk.END, f"{word}: {count} раз(а)\n")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Продвинутый анализатор текста")

    open_button = tk.Button(root, text="Открыть файл", command=open_file)
    open_button.pack(pady=10)

    result_text = tk.Text(root, wrap=tk.WORD, width=40, height=10)
    result_text.pack()

    root.mainloop()
