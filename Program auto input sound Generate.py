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

def close_browser():
    global driver
    if driver:
        driver.quit()
        messagebox.showinfo("Browser Closed", "The browser has been closed.")
        driver = None
    else:
        messagebox.showinfo("No Browser Instance", "There is no open browser instance.")

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
            time.sleep(10)  # Adjust the delay as needed
            
            # Find and click the download button
            download_button = driver.find_element(By.XPATH, "//button[contains(., 'Download')]")
            download_button.click()
            
             # Wait for the download to complete
            time.sleep(20)  # Adjust the delay as needed
            
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

# Create the "Close Browser" button
close_browser_button = tk.Button(root, text="Close Browser", command=close_browser)
close_browser_button.pack()

# Create the "Insert Text from File" button
insert_button = tk.Button(root, text="Insert Text from File", command=insert_text_from_file)
insert_button.pack()

# Create the "Copy Text" button
copy_button = tk.Button(root, text="Copy Text", command=copy_text)
copy_button.pack()

# Start the GUI event loop
root.mainloop()
