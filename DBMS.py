import tkinter as tk
from tkinter import ttk
import pandas as pd
import openpyxl

# Global variables to store data
entry_log_data = []
sales_data = []
inventory_data = []
entry_log_columns = ("ID", "Item", "Company Brand", "Unit", "Price", "Date")
sales_columns = ("Name", "Unit", "Company", "Selling Price", "Date")
inventory_columns = ("Name", "Unit", "Company", "Buying Price", "Selling Price", "Date")

# Function to read data from an Excel file and return as a list of dictionaries
def read_excel(file_name, sheet_name, columns):
    data = []
    try:
        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data.append(dict(zip(columns, row)))
        return data
    except Exception as e:
        print(f"Error reading {sheet_name} from {file_name}: {e}")
        return []

# Function to update the data table based on search query and sorting
def update_data():
    # Update Entry Log tab
    entry_log_tree.delete(*entry_log_tree.get_children())
    entry_log_data = read_excel("DataBase.xlsx", "Entry Log", entry_log_columns)
    for item in entry_log_data:
        entry_log_tree.insert("", "end", values=list(item.values()))

    # Update Sales tab
    sales_tree.delete(*sales_tree.get_children())
    sales_data = read_excel("Sales.xlsx", "Sales", sales_columns)
    for item in sales_data:
        sales_tree.insert("", "end", values=list(item.values()))

    # Update Inventory tab
    inventory_tree.delete(*inventory_tree.get_children())
    inventory_data = read_excel("Inventory.xlsx", "Inventory", inventory_columns)
    for item in inventory_data:
        inventory_tree.insert("", "end", values=list(item.values()))

    # Get the search query from the entry field
    search_query = search_entry.get()





# Function to open the data entry script
def open_data_entry():
    from tkinter import messagebox
    import datetime

    df = pd.read_excel("DataBase.xlsx")

    # Function to handle the button click and save data to the database
    def save_data():
        item = item_entry.get()
        company_brand = company_brand_var.get()
        unit = unit_entry.get()
        price = "RM" + str("{:.2f}".format(float(price_entry.get())))
        date = datetime.datetime.now()  # Retrieve the date from the DateEntry widget
        max_id = df["ID"].max()  # Get the maximum ID value
        new_id = max_id + 1 if not pd.isna(max_id) else 1  # Calculate the new ID

        data = {"ID": new_id, "Item": item, "Company Brand": company_brand, "Unit": unit, "Price": price, "Date": date}

        # You can add code here to save the data to your database
        df_to_append = pd.DataFrame([data], columns=["ID", "Item", "Company Brand", "Unit", "Price", "Date"])

        # Load the existing Excel file
        excel_file_path = "DataBase.xlsx"
        try:
            df_existing = pd.read_excel(excel_file_path, sheet_name="Entry log")
        except FileNotFoundError:
            df_existing = pd.DataFrame(columns=["ID", "Item", "Company Brand", "Unit", "Price", "Date"])

        # Concatenate the existing and new data
        df_updated = pd.concat([df_existing, df_to_append], ignore_index=True)

        # Save the updated data to the Excel file
        df_updated.to_excel(excel_file_path, index=False, engine='openpyxl')

        # Display a message box to confirm the data is saved
        update_data()
        messagebox.showinfo("Success", "Data saved successfully!")

    # Create the main window
    root = tk.Tk()
    root.title("Database Input System")

    # Create and arrange widgets
    item_label = ttk.Label(root, text="Item:")
    item_entry = ttk.Entry(root)

    company_brand_label = ttk.Label(root, text="Company Brand:")
    company_brand_var = ttk.Entry(root)
    #company_brands = companylist
    #company_brand_var = tk.StringVar()
    #company_brand_dropdown = ttk.Combobox(root, textvariable=company_brand_var, values=company_brands)

    unit_label = ttk.Label(root, text="Unit:")
    unit_entry = ttk.Entry(root)

    price_label = ttk.Label(root, text="Price:")
    price_entry = ttk.Entry(root)

    save_button = ttk.Button(root, text="Save", command=save_data)

    # Arrange widgets using grid layout
    item_label.grid(row=0, column=0, padx=10, pady=5)
    item_entry.grid(row=0, column=1, padx=10, pady=5)

    company_brand_label.grid(row=1, column=0, padx=10, pady=5)
    company_brand_var.grid(row=1, column=1, padx=10, pady=5)

    unit_label.grid(row=2, column=0, padx=10, pady=5)
    unit_entry.grid(row=2, column=1, padx=10, pady=5)

    price_label.grid(row=3, column=0, padx=10, pady=5)
    price_entry.grid(row=3, column=1, padx=10, pady=5)

    save_button.grid(row=5, columnspan=2, padx=10, pady=10)

    # Start the main loop
    root.mainloop()


if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    root.title("Data Table")

    # Create a Frame to hold the search bar, sorting, and button
    search_frame = ttk.Frame(root)
    search_frame.pack()

    # Create an Entry widget for search and place it in the search frame
    search_label = ttk.Label(search_frame, text="Search:")
    search_label.pack(side=tk.LEFT)
    search_entry = ttk.Entry(search_frame)
    search_entry.pack(side=tk.LEFT)

    # Create a Combobox for sorting column selection
    sort_column_var = tk.StringVar()
    sort_column_label = ttk.Label(search_frame, text="Sort by:")
    sort_column_label.pack(side=tk.LEFT)
    # Include "Not Sorting" option in the values
    sort_column_combobox = ttk.Combobox(search_frame, textvariable=sort_column_var,
                                        values=(" ", "ID", "Item", "Company Brand", "Unit", "Price", "Date"))
    sort_column_combobox.pack(side=tk.LEFT)
    sort_column_combobox.set(" ")  # Set a default sorting option

    # Create a Combobox for sorting order selection
    sort_order_var = tk.StringVar()
    sort_order_label = ttk.Label(search_frame, text="Order:")
    sort_order_label.pack(side=tk.LEFT)
    sort_order_combobox = ttk.Combobox(search_frame, textvariable=sort_order_var, values=("Ascending", "Descending"))
    sort_order_combobox.pack(side=tk.LEFT)
    sort_order_combobox.set("Ascending")  # Set a default sorting order

    # Create a Search button to trigger data search and sorting
    search_button = ttk.Button(search_frame, text="Search & Sort", command=update_data)
    search_button.pack(side=tk.LEFT)

    # Create a ttk.Style object to customize the tab style
    style = ttk.Style()

    # Increase font size and make tabs bold
    style.configure("TNotebook.Tab", font=("Helvetica", 9, "bold"), padding=[10, 1])

    # Create a notebook (tabbed interface)
    notebook = ttk.Notebook(root)

    # Create the Entry Log tab
    entry_log_frame = ttk.Frame(notebook)
    entry_log_data = read_excel("DataBase.xlsx", "Entry Log", entry_log_columns)
    entry_log_tree = ttk.Treeview(entry_log_frame, columns=entry_log_columns, show="headings")
    for col in entry_log_columns:
        entry_log_tree.heading(col, text=col)
    for item in entry_log_data:
        entry_log_tree.insert("", "end", values=list(item.values()))
    entry_log_tree.pack(fill="both", expand=True)
    notebook.add(entry_log_frame, text="Entry Log")

    # Create the Sales tab
    sales_frame = ttk.Frame(notebook)
    sales_data = read_excel("Sales.xlsx", "Sales", sales_columns)
    sales_tree = ttk.Treeview(sales_frame, columns=sales_columns, show="headings")
    for col in sales_columns:
        sales_tree.heading(col, text=col)
    for item in sales_data:
        sales_tree.insert("", "end", values=list(item.values()))
    sales_tree.pack(fill="both", expand=True)
    notebook.add(sales_frame, text="Sales")

    # Create the Inventory tab
    inventory_frame = ttk.Frame(notebook)
    inventory_data = read_excel("Inventory.xlsx", "Inventory", inventory_columns)
    inventory_tree = ttk.Treeview(inventory_frame, columns=inventory_columns, show="headings")
    for col in inventory_columns:
        inventory_tree.heading(col, text=col)
    for item in inventory_data:
        inventory_tree.insert("", "end", values=list(item.values()))
    inventory_tree.pack(fill="both", expand=True)
    notebook.add(inventory_frame, text="Inventory")

    # Pack the notebook
    notebook.pack(fill="both", expand=True)

    # Call the update_data function to display all data on startup
    update_data()



    # Start the main loop
    root.mainloop()
