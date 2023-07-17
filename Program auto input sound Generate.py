import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import os
import docx
import PyPDF2
import webbrowser

def open_mic_monster():
    url = "https://micmonster.com/"
    try:
        webbrowser.open(url)
        messagebox.showinfo("MicMonster Opened", "MicMonster website has been opened in the default web browser.")
    except webbrowser.Error:
        messagebox.showerror("Error", "Failed to open MicMonster website.")

def insert_text_from_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("Word Documents", "*.docx"), ("PDF Files", "*.pdf")])
    if file_path:
        extension = os.path.splitext(file_path)[1].lower()
        try:
            if extension == ".txt":
                with open(file_path, "r") as file:
                    text = file.read()
                    text_box.delete("1.0", tk.END)  # Clear the existing text
                    text_box.insert(tk.END, text)  # Insert the file contents into the text box
                    messagebox.showinfo("File Read", "Text file contents have been inserted into the text box.")
            elif extension == ".docx":
                doc = docx.Document(file_path)
                paragraphs = [paragraph.text for paragraph in doc.paragraphs]
                text = "\n".join(paragraphs)
                text_box.delete("1.0", tk.END)  # Clear the existing text
                text_box.insert(tk.END, text)  # Insert the file contents into the text box
                messagebox.showinfo("File Read", "DOCX file contents have been inserted into the text box.")
            elif extension == ".pdf":
                with open(file_path, "rb") as file:
                    pdf = PyPDF2.PdfReader(file)
                    pages = [page.extract_text() for page in pdf.pages]
                    text = "\n".join(pages)
                    text_box.delete("1.0", tk.END)  # Clear the existing text
                    text_box.insert(tk.END, text)  # Insert the file contents into the text box
                    messagebox.showinfo("File Read", "PDF file contents have been inserted into the text box.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("No File Selected", "No file was selected.")

def copy_text():
    text = text_box.get("1.0", tk.END)
    if text.strip() == "":
        messagebox.showwarning("Empty Text", "There is no text to copy.")
    else:
        root.clipboard_clear()
        root.clipboard_append(text)
        messagebox.showinfo("Text Copied", "The text has been copied to the clipboard.")
        
def insert_text_to_web():
    text = text_box.get("1.0", tk.END).strip()  # Get the text from the text box
    try:
        webbrowser.open_new_tab(f"https://micmonster.com/?text={text}")
        messagebox.showinfo("Text Inserted", "The text has been inserted into the web.")
    except:
        messagebox.showerror("Error", "Failed to insert text into the web.")

# Create the main window
root = tk.Tk()
root.title("Text Action")
root.geometry("1280x720")

# Create the input text box
text_box = tk.Text(root, height=30)  # Adjust the height as per your requirement
text_box.pack(pady=10)

# Create the "Insert Text from File" button
insert_button = tk.Button(root, text="Insert Text from File", command=insert_text_from_file)
insert_button.pack()

# Create the "Copy Text" button
copy_button = tk.Button(root, text="Copy Text", command=copy_text)
copy_button.pack()

# Create the "Open MicMonster" button
mic_monster_button = tk.Button(root, text="Open MicMonster", command=open_mic_monster)
mic_monster_button.pack()

# Create the "Insert Text to Web" button
insert_web_button = tk.Button(root, text="Insert Text to Web", command=insert_text_to_web)
insert_web_button.pack()

# Start the GUI event loop
root.mainloop()
