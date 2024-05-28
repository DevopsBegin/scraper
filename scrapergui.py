import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
from scraper import scrape_from_html, save_to_log, save_to_excel

def scrape_and_save():
    try:
        # URL of amsall.html
        html_url = 'http://megds55hi0.asiapacific.hpqcorp.net/AMSEmail/AMSAll.html'

        # Download amsall.html
        response = requests.get(html_url)
        response.raise_for_status()  # Raise an HTTPError for bad responses

        # Extract data from amsall.html
        email_data = scrape_from_html(response.content)

        # Save data to log file
        save_to_log(email_data)

        # Ask user to choose where to save Excel file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            # Save data to Excel file
            save_to_excel(email_data, file_path)
            status_label.config(text="Data saved to Excel file successfully.", foreground="green")
        else:
            status_label.config(text="No file selected.", foreground="orange")
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Error", f"Failed to retrieve URL: {e}")
        status_label.config(text="Failed to retrieve URL.", foreground="red")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label.config(text="An error occurred.", foreground="red")

def exit_application():
    root.quit()

# Create the main window
root = tk.Tk()
root.title("Scraper Application")
root.geometry("400x200")

# Set a modern theme if available
style = ttk.Style()
if "vista" in style.theme_names():
    style.theme_use('vista')

# Create a frame for the buttons
button_frame = ttk.Frame(root, padding="20")
button_frame.pack(expand=True)

# Create and place the scrape button
scrape_button = ttk.Button(button_frame, text="Scrape and Save", command=scrape_and_save, width=20)
scrape_button.grid(row=0, column=0, pady=10)

# Create and place the exit button
exit_button = ttk.Button(button_frame, text="Exit", command=exit_application, width=20)
exit_button.grid(row=1, column=0, pady=10)

# Create a status bar
status_label = ttk.Label(root, text="Ready", relief=tk.SUNKEN, anchor=tk.W, padding=(10, 2))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# Run the GUI event loop
root.mainloop()
