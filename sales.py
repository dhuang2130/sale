import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def convert_xls_to_xlsx(file_path):
    # Convert .xls file to .xlsx
    xls = pd.ExcelFile(file_path)
    xlsx_file_path = file_path.replace('.xls', '.xlsx')

    # Save it as .xlsx
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return xlsx_file_path

def process_file(file_path):
    try:
        # Check if the file is .xls and convert to .xlsx
        if file_path.endswith('.xls'):
            file_path = convert_xls_to_xlsx(file_path)

        # Load the converted .xlsx file
        sales_data_df = pd.read_excel(file_path, sheet_name='Sales Data', header=1)
        product_key_df = pd.read_excel(file_path, sheet_name='Product Key', header=None)

        # Rename columns for easier access
        product_key_df.columns = ['Product Key', 'Product Name', 'Price']

        # Initialize an empty dictionary to store the count of items sold per month
        product_sales = {product_key: {month: 0 for month in range(1, 13)} for product_key in product_key_df['Product Key']}

        # Process each row in the sales data
        for _, row in sales_data_df.iterrows():
            date = pd.to_datetime(row['Order Date'])
            month = date.month
            items = row['Purchase'].split()

            for i in range(0, len(items), 2):
                count = int(items[i])
                product_key = items[i + 1]

                if product_key in product_sales:
                    product_sales[product_key][month] += count

        # Prepare the data to be written into the new sheet
        output_data = []
        for _, row in product_key_df.iterrows():
            product_key = row['Product Key']
            product_name = row['Product Name']
            price = row['Price']
            monthly_sales = [product_sales[product_key][month] for month in range(1, 13)]
            output_data.append([product_key, product_name, price] + monthly_sales)

        # Create a DataFrame for the output data
        columns = ['Product Key', 'Product Name', 'Price($)'] + [pd.to_datetime(f'{month}/1/2024').strftime('%b') for month in range(1, 13)]
        output_df = pd.DataFrame(output_data, columns=columns)

        # Write the output data to a new sheet in the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            output_df.to_excel(writer, sheet_name='Monthly Sales', index=False)

        messagebox.showinfo("Success", "Monthly sales data has been successfully written to the 'Monthly Sales' sheet.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        file_label.config(text=f"Selected file: {file_path}")
        go_button.config(state=tk.NORMAL)  # Enable the Go button

def go_button_clicked():
    file_path = file_label.cget("text").replace("Selected file: ", "")
    if file_path:
        process_file(file_path)

# Create the main window
root = tk.Tk()
root.title("Excel Sales Data Processor")
root.geometry("500x200")  # Set the default size

# Create and place the "Select File" button
select_button = tk.Button(root, text="Select Excel File", command=select_file)
select_button.pack(pady=10)

# Create a label to display the selected file name
file_label = tk.Label(root, text="No file selected")
file_label.pack(pady=10)

# Create and place the "Go" button
go_button = tk.Button(root, text="Go", command=go_button_clicked, state=tk.DISABLED)
go_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
