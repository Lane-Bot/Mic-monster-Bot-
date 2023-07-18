import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import os
import docx
import PyPDF2
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = None  # Global variable to hold the browser instance
processed_files = set()  # Set to store the names of processed files

def open_mic_monster():
    global driver
    url = "https://micmonster.com/"
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
    folder_path = filedialog.askdirectory()
    if folder_path:
        files = get_files_in_folder(folder_path)
        for file_path in files:
            if file_path not in processed_files:
                process_file(file_path)
                processed_files.add(file_path)
    else:
        messagebox.showwarning("No Folder Selected", "No folder was selected.")

def get_files_in_folder(folder_path):
    files = []
    for root, dirs, filenames in os.walk(folder_path):
        for filename in filenames:
            files.append(os.path.join(root, filename))
    return files

def process_file(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    try:
        if extension == ".txt":
            with open(file_path, "r") as file:
                text = file.read()
                text_box.delete("1.0", tk.END)  # Clear the existing text
                text_box.insert(tk.END, text)  # Insert the file contents into the text box
                messagebox.showinfo("File Read", "Text file contents have been inserted into the text box.")
                insert_text_into_website(text)  # Insert the text into the website
        elif extension == ".docx":
            doc = docx.Document(file_path)
            paragraphs = [paragraph.text for paragraph in doc.paragraphs]
            text = "\n".join(paragraphs)
            text_box.delete("1.0", tk.END)  # Clear the existing text
            text_box.insert(tk.END, text)  # Insert the file contents into the text box
            messagebox.showinfo("File Read", "DOCX file contents have been inserted into the text box.")
            insert_text_into_website(text)  # Insert the text into the website
        elif extension == ".pdf":
            with open(file_path, "rb") as file:
                pdf = PyPDF2.PdfReader(file)
                pages = [page.extract_text() for page in pdf.pages]
                text = "\n".join(pages)
                text_box.delete("1.0", tk.END)  # Clear the existing text
                text_box.insert(tk.END, text)  # Insert the file contents into the text box
                messagebox.showinfo("File Read", "PDF file contents have been inserted into the text box.")
                insert_text_into_website(text)  # Insert the text into the website
    except Exception as e:
        messagebox.showerror("Error", str(e))

def copy_text():
    text = text_box.get("1.0", tk.END)
    if text.strip() == "":
        messagebox.showwarning("Empty Text", "There is no text to copy.")
    else:
        root.clipboard_clear()
        root.clipboard_append(text)
        messagebox.showinfo("Text Copied", "The text has been copied to the clipboard.")


def insert_text_into_website(text):
    if text == "":
        messagebox.showwarning("Empty Text", "Please enter some text.")
    elif not driver:
        messagebox.showwarning("No Browser", "Please open the browser first.")
    else:
        try:
            text_area = driver.find_element(By.CSS_SELECTOR, ".text-area")
            text_area.clear()  # Clear any existing text
            
            # Introduce a small delay to allow the web page to load properly
            time.sleep(1)
            
            text_area.send_keys(text)  # Insert the text into the web text box
            
            # Wait for the "Generating..." span to disappear
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.XPATH, "//span[text()='Generating...']"))
            )
            
            # Wait for the "Generate" button to become clickable and visible
            generate_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[@class='primary-text' and text()='Generate']"))
            )
            
            # Click on the "Generate" button
            generate_button.click()
            
            # Wait for the page to refresh
            page_refresh_delay = int(page_refresh_entry.get())
            time.sleep(page_refresh_delay)  # Adjust the delay as needed
            
            # Find and click the download button
            download_button = driver.find_element(By.XPATH, "//button[contains(., 'Download')]")
            download_button.click()
            
            # Wait for the download to complete
            download_delay = int(download_delay_entry.get())
            time.sleep(download_delay)  # Adjust the delay as needed
            
            # Refresh the page
            driver.refresh()
            
            messagebox.showinfo("Text Inserted", "The text has been inserted into the website.")
        except Exception as e:
            print("Exception:", str(e))
            messagebox.showerror("Error", "Failed to insert text into the website.")

# Create the main window
root = tk.Tk()
root.title("Mic Monster")
root.geometry("1280x720")

# Create the input text box
text_box = tk.Text(root, height=30)  # Adjust the height as per your requirement
text_box.pack(pady=30)

# Create the "Open MicMonster" button
mic_monster_button = tk.Button(root, text="Open MicMonster", command=open_mic_monster)
mic_monster_button.pack()

# Create the "Insert Text from File" button
insert_button = tk.Button(root, text="Insert Text from Folder", command=insert_text_from_file)
insert_button.pack()

# Create the "Copy Text" button
copy_button = tk.Button(root, text="Copy Text", command=copy_text)
copy_button.pack()

# Create a frame for the sleep durations
sleep_frame = tk.Frame(root)
sleep_frame.pack()

# Create the label and entry for page refresh delay
page_refresh_label = tk.Label(sleep_frame, text="Page Refresh Delay (seconds):")
page_refresh_label.pack(side=tk.LEFT)
page_refresh_entry = tk.Entry(sleep_frame)
page_refresh_entry.pack(side=tk.LEFT)
page_refresh_entry.insert(tk.END, "10")  # Set a default value

# Create the label and entry for download delay
download_delay_label = tk.Label(sleep_frame, text="Download Delay (seconds):")
download_delay_label.pack(side=tk.LEFT)
download_delay_entry = tk.Entry(sleep_frame)
download_delay_entry.pack(side=tk.LEFT)
download_delay_entry.insert(tk.END, "10")  # Set a default value

# Start the GUI event loop
root.mainloop()
