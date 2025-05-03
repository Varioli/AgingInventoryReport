import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

# --- Color Theme for the GUI ---
BG_COLOR = "#0A1128"
FRAME_BG = "#001F54"
TEXT_COLOR = "#FEFCFB"
ACCENT_COLOR = "#1282A2"
BUTTON_BG = "#034078"
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")

# Global variables to hold file paths after upload
selected_inventory_file = None
selected_sales_file = None

# --- Helper Functions ---

# Cleans the product description to make it easier to match across reports
def clean_description(text):
    if pd.isna(text):
        return ''
    return str(text).lower().replace(" ", "").replace("~", "").replace("-", "").replace("/", "").replace("tl", "")

# Classifies tires based on how long it's been since they were last sold
def classify_status(days_idle):
    if pd.isna(days_idle):
        return 'NO SALES HISTORY'
    elif days_idle >= 90:
        return 'EXTREME – 90+ DAYS'
    elif days_idle >= 60:
        return 'WARNING – 60+ DAYS'
    else:
        return 'ACTIVE'

# Core function that processes inventory and sales data
# Merges the reports, calculates days since last sale, assigns status, and generates a temporary Excel file
def process_inventory_and_sales(inventory_file, sales_file):
    try:
        # Load and clean inventory data
        inv_df = pd.read_csv(inventory_file, skiprows=3)
        inv_df = inv_df[['SkuCode', 'Description', 'QuantityOnHand']]
        inv_df['CleanDesc'] = inv_df['Description'].apply(clean_description)

        # Load and clean sales data
        sales_df = pd.read_csv(sales_file, skiprows=3)
        sales_df = sales_df[['Description', 'DateCompleted']]
        sales_df['CleanDesc'] = sales_df['Description'].apply(clean_description)
        sales_df['DateCompleted'] = pd.to_datetime(sales_df['DateCompleted'], errors='coerce')

        # Get most recent sale date for each cleaned description
        last_sales = sales_df.groupby('CleanDesc')['DateCompleted'].max().reset_index()
        last_sales.columns = ['CleanDesc', 'LastSoldDate']

        # Merge with inventory and calculate days idle
        today = pd.to_datetime(datetime.today().date())
        merged = inv_df.merge(last_sales, on='CleanDesc', how='left')
        merged['DaysIdle'] = (today - merged['LastSoldDate']).dt.days
        merged['Status'] = merged['DaysIdle'].apply(classify_status)
        merged.sort_values(by='DaysIdle', ascending=False, inplace=True)

        # Output report as temporary file
        temp_output = f"AgingInventory_TEMP.xlsx"
        merged[['SkuCode', 'Description', 'QuantityOnHand', 'LastSoldDate', 'DaysIdle', 'Status']].to_excel(temp_output, index=False)
        return temp_output
    except Exception as e:
        return str(e)

# --- GUI Setup ---
def run_gui():
    # Prompts user to upload the inventory report file
    def upload_inventory():
        global selected_inventory_file
        file = filedialog.askopenfilename(title="Select Inventory Report CSV", filetypes=[("CSV files", "*.csv")])
        if file:
            selected_inventory_file = file
            inv_label.config(text=f"Inventory: {os.path.basename(file)}")

    # Prompts user to upload the sales report file — blocked if inventory isn't uploaded first
    def upload_sales():
        global selected_sales_file
        if not selected_inventory_file:
            messagebox.showwarning("Upload Inventory First", "Please upload the Inventory file before the Sales file.")
            return
        file = filedialog.askopenfilename(title="Select Sales Report CSV", filetypes=[("CSV files", "*.csv")])
        if file:
            selected_sales_file = file
            sales_label.config(text=f"Sales: {os.path.basename(file)}")

    # Resets both file paths and GUI labels
    def clear_files():
        global selected_inventory_file, selected_sales_file
        selected_inventory_file = None
        selected_sales_file = None
        inv_label.config(text="Inventory: Not selected")
        sales_label.config(text="Sales: Not selected")

    # Generates the final aging report after both files are uploaded
    def run_report():
        if not selected_inventory_file or not selected_sales_file:
            messagebox.showerror("Missing Files", "Please upload both Inventory and Sales CSV files.")
            return

        today = datetime.today().strftime("%Y_%m_%d")
        default_name = f"AgingInventory_Report_{today}.xlsx"

        # Prompt user to save the final output
        save_path = filedialog.asksaveasfilename(
            title="Save Aging Report As",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return  # Cancelled

        result = process_inventory_and_sales(selected_inventory_file, selected_sales_file)

        # If the result isn't a valid .xlsx file, show error
        if isinstance(result, str) and not result.endswith(".xlsx"):
            messagebox.showerror("Error", f"Failed to generate report:{result}")
            return

        # Attempt to move the temp file to the chosen location
        try:
            os.replace(result, save_path)
            messagebox.showinfo("Report Created", f"Report successfully saved as:{save_path}")
        except Exception as e:
            messagebox.showerror("File Error", f"Could not save report:{e}")

    # closes out splash screen and loads main window
    def launch_main():
        splash.destroy()
        build_main_window()

    # Builds the main GUI window with all buttons and labels
    def build_main_window():
        global inv_label, sales_label

        root = tk.Tk()
        root.title("Tire Aging Report Generator - By Josh Clark")
        root.configure(bg=BG_COLOR)
        root.geometry("600x360")
        root.resizable(False, False)

        tk.Label(root, text="Tire Aging Report Generator", font=("Segoe UI", 16, "bold"), fg=TEXT_COLOR, bg=BG_COLOR).pack(pady=10)
        tk.Label(root, text="Upload Inventory FIRST, then Sales CSV", font=FONT, fg=ACCENT_COLOR, bg=BG_COLOR).pack(pady=2)

        # Frame to display selected file names
        display_frame = tk.Frame(root, bg=FRAME_BG)
        display_frame.pack(pady=10, padx=15, fill="both", expand=True)

        inv_label = tk.Label(display_frame, text="Inventory: Not selected", font=FONT, fg=TEXT_COLOR, bg=FRAME_BG, anchor="w", justify="left", wraplength=500)
        inv_label.pack(padx=10, pady=5, fill="x")

        sales_label = tk.Label(display_frame, text="Sales: Not selected", font=FONT, fg=TEXT_COLOR, bg=FRAME_BG, anchor="w", justify="left", wraplength=500)
        sales_label.pack(padx=10, pady=5, fill="x")

        # Frame for action buttons
        button_frame = tk.Frame(root, bg=BG_COLOR)
        button_frame.pack(pady=10)

        # Upload, Clear, and Run buttons
        tk.Button(button_frame, text="Upload Inventory", font=FONT_BOLD, bg=BUTTON_BG, fg=TEXT_COLOR,
                  activebackground=ACCENT_COLOR, width=18, command=upload_inventory).grid(row=0, column=0, padx=10)

        tk.Button(button_frame, text="Upload Sales", font=FONT_BOLD, bg=BUTTON_BG, fg=TEXT_COLOR,
                  activebackground=ACCENT_COLOR, width=18, command=upload_sales).grid(row=0, column=1, padx=10)

        tk.Button(button_frame, text="Clear Uploads", font=FONT_BOLD, bg="#444", fg=TEXT_COLOR,
                  activebackground="#666", width=18, command=clear_files).grid(row=1, column=0, padx=10, pady=10)

        tk.Button(button_frame, text="Run Report", font=FONT_BOLD, bg="#1282A2", fg="white",
                  activebackground=ACCENT_COLOR, width=18, command=run_report).grid(row=1, column=1, padx=10, pady=10)

        # Footer
        tk.Label(root, text="© 2025", font=("Segoe UI", 8), fg=ACCENT_COLOR, bg=BG_COLOR).pack(side="bottom", pady=10)
        root.mainloop()

    # --- Splash Screen ---
    splash = tk.Tk()
    splash.overrideredirect(True)  # Removes window border
    splash.configure(bg=FRAME_BG)
    splash.geometry("400x200+500+300")

    # Splash content
    tk.Label(splash, text="INVENTORY AGING REPORTER", font=("Segoe UI", 14, "bold"),
             fg=TEXT_COLOR, bg=FRAME_BG).pack(pady=(35, 5))
    tk.Label(splash, text="by Josh Clark", font=("Segoe UI", 10), fg=ACCENT_COLOR, bg=FRAME_BG).pack()
    tk.Label(splash, text="loading...", font=("Segoe UI", 7), fg=ACCENT_COLOR, bg=FRAME_BG).pack()
    tk.Label(splash, text="~kaizen~", font=("Segoe UI", 9, "italic"), fg=TEXT_COLOR, bg=FRAME_BG).pack(pady=(40, 0))

    splash.after(3500, launch_main)
    splash.mainloop()

if __name__ == "__main__":
    try:
        run_gui()
    except KeyboardInterrupt:
        print("Exited manually.")
