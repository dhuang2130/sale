import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def replace_product_names(sku_file_path, output_file_path):
    # Load the SKU file
    sku_df = pd.read_excel(sku_file_path, sheet_name=None)

    # Create a dictionary from the 'key' sheet
    key_df = sku_df['key']
    key_dict = dict(zip(key_df.iloc[:, 1], key_df.iloc[:, 0]))

    # Sort the keys by length in descending order to prioritize longer keys
    sorted_key_dict = {k: v for k, v in sorted(key_dict.items(), key=lambda item: len(item[0]), reverse=True)}

    # Replace product names in 'Purchase' column of 'Sheet1' with the corresponding codes
    sheet1_df = sku_df['Sheet1'].copy()

    def replace_product_name(purchase_str):
        if isinstance(purchase_str, str):
            # Remove unwanted characters
            purchase_str = purchase_str.replace('Ã—', '').replace(',', '')
            # Replace product names with codes
            for product_name, product_code in sorted_key_dict.items():
                if product_name in purchase_str:
                    purchase_str = purchase_str.replace(product_name, product_code)
        return purchase_str

    sheet1_df['Purchase'] = sheet1_df['Purchase'].apply(replace_product_name)

    # Ensure that the date column is correctly formatted as dates
    sheet1_df['Order Date'] = pd.to_datetime(sheet1_df['Order Date'], errors='coerce')

    # Save the updated 'Sheet1' back to the Excel file
    with pd.ExcelWriter(output_file_path, date_format='YYYY-MM-DD') as writer:
        sheet1_df.to_excel(writer, index=False, sheet_name='Sheet1')
        key_df.to_excel(writer, index=False, sheet_name='key')

def select_input_file():
    input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_file_path)

def select_output_file():
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_file_path)

def run_replacement():
    input_file = input_entry.get()
    output_file = output_entry.get()
    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return
    try:
        replace_product_names(input_file, output_file)
        messagebox.showinfo("Success", "Product names replaced successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main application window
root = tk.Tk()
root.title("SKU Replacement Tool")

# Input file selection
tk.Label(root, text="Select input Excel file:").grid(row=0, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=10, pady=10)
input_button = tk.Button(root, text="Browse...", command=select_input_file)
input_button.grid(row=0, column=2, padx=10, pady=10)

# Output file selection
tk.Label(root, text="Select output Excel file:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
output_button = tk.Button(root, text="Browse...", command=select_output_file)
output_button.grid(row=1, column=2, padx=10, pady=10)

# Run button
run_button = tk.Button(root, text="Run Replacement", command=run_replacement)
run_button.grid(row=2, column=0, columnspan=3, pady=20)

# Start the Tkinter event loop
root.mainloop()
