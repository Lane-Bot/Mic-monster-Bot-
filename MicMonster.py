import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter.filedialog as filedialog
import os
import docx
import PyPDF2
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

driver = None  # Global variable to hold the browser instance
file_list = []  # List to store the files in the selected folder
processed_files = set()  # Set to store the names of processed files
current_file = ""  # Global variable to store the name of the current file being processed
cancel_processing = False  # Flag to indicate if the file processing should be canceled

def open_mic_monster():
    global driver
    url = "https://app.micmonster.com/"
    try:
        options = webdriver.ChromeOptions()
        options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"  # Path to Chrome executable
        driver = webdriver.Chrome(options=options)  # Let chromedriver automatically locate the executable
        driver.get(url)
        messagebox.showinfo("MicMonster Opened", "MicMonster website has been opened in Google Chrome.")
    except:
        messagebox.showerror("Error", "Failed to open MicMonster website.")
        return

def insert_text_from_file():
    global cancel_processing, file_list
    folder_path = filedialog.askdirectory()
    if folder_path:
        file_list = get_files_in_folder(folder_path)
        progress_bar["maximum"] = len(file_list)
        cancel_processing = False
        process_next_file()
    else:
        messagebox.showwarning("No Folder Selected", "No folder was selected.")

def get_files_in_folder(folder_path):
    files = []
    for root, dirs, filenames in os.walk(folder_path):
        for filename in filenames:
            files.append((os.path.join(root, filename), filename))  # Store file path and name as tuple
    return files

def process_next_file():
    global cancel_processing, file_list
    if cancel_processing or not file_list:
        # Cancel or no more files to process
        progress_bar["value"] = 0
        current_file_label.config(text="Processing: None")
        messagebox.showinfo("Processing Complete", "All files have been processed.")
        return
    
    file_path, file_name = file_list.pop(0)  # Get the next file path and file name from the list
    
    try:
        extension = os.path.splitext(file_path)[1].lower()
        if extension == ".txt":
            with open(file_path, "r") as file:
                text = file.read()
                name_text_box.delete(1.0, tk.END)
                name_text_box.insert(tk.END, remove_extension(file_name))  # Update name_text_box with the file name (without extension)
                insert_text_into_website(text, file_name)  # Insert the text into the website with the file name
        elif extension == ".docx":
            doc = docx.Document(file_path)
            paragraphs = [paragraph.text for paragraph in doc.paragraphs]
            text = "\n".join(paragraphs)
            name_text_box.delete(1.0, tk.END)
            name_text_box.insert(tk.END, remove_extension(file_name))  # Update name_text_box with the file name (without extension)
            insert_text_into_website(text, file_name)  # Insert the text into the website with the file name
        elif extension == ".pdf":
            with open(file_path, "rb") as file:
                pdf = PyPDF2.PdfReader(file)
                pages = [page.extract_text() for page in pdf.pages]
                text = "\n".join(pages)
                name_text_box.delete(1.0, tk.END)
                name_text_box.insert(tk.END, remove_extension(file_name))  # Update name_text_box with the file name (without extension)
                insert_text_into_website(text, file_name)  # Insert the text into the website with the file name
                
        processed_files.add(file_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))
    
    # Update progress and current file label
    progress_bar["value"] = len(processed_files) + 1
    current_file_label.config(text=f"Processing: {file_name}")
    
    # Schedule next file processing
    root.after(10000, process_next_file)

def insert_text_into_website(text, file_name):
    if text == "":
        messagebox.showwarning("Empty Text", "Please enter some text.")
    elif not driver:
        messagebox.showwarning("No Browser", "Please open the browser first.")
    else:
        try:
            # Insert the file name (without extension) into the name text box on the website
            driver.execute_script("document.getElementById('user-voice-name').value = arguments[0];", remove_extension(file_name))

            text_area = driver.find_element(By.CSS_SELECTOR, "#current-content")
            text_area.clear()  # Clear any existing text

            # Introduce a small delay to allow the web page to load properly
            time.sleep(1)

            # Execute JavaScript to insert the text into the web text box
            driver.execute_script("arguments[0].value = arguments[1];", text_area, text)

            # Find and click the "CREATE" button
            create_button = driver.find_element(By.CSS_SELECTOR, ".btn.btn-special.btn-md.mt-3")
            create_button.click()
            
            # Wait for the create process to finish (adjust the delay as needed)
            time.sleep(10)

        except NoSuchElementException as e:
            print("Element Not Found:", str(e))
            messagebox.showerror("Error", "Failed to locate elements on the website.")
        except Exception as e:
            print("Exception:", str(e))
            messagebox.showerror("Error", "Failed to insert text into the website.")

def cancel_process():
    global cancel_processing
    cancel_processing = True

def remove_extension(file_name):
    return os.path.splitext(file_name)[0]

# Create the main window
root = tk.Tk()
root.title("Mic Monster")
root.geometry("1280x820")

# Create the label for the name
name_label = tk.Label(root, text="NAME")
name_label.pack(pady=20)

# Create the name text box
name_text_box = tk.Text(root, height=1)
name_text_box.pack()

# Create the label for the TEXTBOX
name_label = tk.Label(root, text="TEXT BOX")
name_label.pack(pady=10)
# Create the input text box
text_box = tk.Text(root, height=30)  # Adjust the height as per your requirement
text_box.pack(pady=10)

# Create the "Open MicMonster" button
mic_monster_button = tk.Button(root, text="Open MicMonster", command=open_mic_monster)
mic_monster_button.pack()

# Create the "Insert Text from File" button
insert_button = tk.Button(root, text="Insert Text from Folder", command=insert_text_from_file)
insert_button.pack()

# Create a frame for the progress bar
progress_frame = tk.Frame(root)
progress_frame.pack()

# Create the progress bar
progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
progress_bar.pack(fill=tk.X, padx=10, pady=10)

# Create a frame for the current file label and cancel button
file_frame = tk.Frame(root)
file_frame.pack()

# Create the label for the current file
current_file_label = tk.Label(file_frame, text="Processing: None")
current_file_label.pack(side=tk.LEFT)

# Create the cancel button
cancel_button = tk.Button(file_frame, text="Cancel", command=cancel_process)
cancel_button.pack(side=tk.LEFT)


# Start the GUI event loop
root.mainloop()
