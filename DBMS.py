import tkinter as tk
from tkinter import ttk
import pandas as pd
import openpyxl
from tkinter import messagebox
import datetime


# Global variables to store data
entry_log_data = []
sales_data = []
inventory_data = []
entry_log_columns = ("ID", "Item", "Company Brand", "Unit", "Price", "Date")
sales_columns = ("Name", "Unit","Company",  "Selling Price", "Date")
inventory_columns = ("Name", "Unit","Company",  "Buying Price", "Selling Price", "Date")
inventory_excel_file_path="Inventory.xlsx"
inventory_sheet_name="Inventory"

def update_sort_column_combobox():
    # Get the currently selected tab
    current_tab = notebook.index(notebook.select())

    # Update the sort_column_combobox values based on the current tab
    if current_tab == 0:  # Entry Log tab
        sort_column_combobox['values'] = (" ", "ID", "Item", "Company Brand", "Price", "Date")
    elif current_tab == 1:  # Sales tab
        sort_column_combobox['values'] = (" ", "Name", "Company", "Selling Price", "Date")
    elif current_tab == 2:  # Inventory tab
        sort_column_combobox['values'] = (" ", "Name", "Company", "Buying Price", "Selling Price", "Date")
    else:
        sort_column_combobox['values'] = (" ")  # Default if no tab is selected


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
    # Get the search query from the entry field
    search_query = search_entry.get().lower()

    # Get the currently selected tab
    current_tab = notebook.index(notebook.select())

    # Get the selected sorting column and order
    sort_column = sort_column_var.get()
    sort_order = sort_order_var.get()

    # Read data from Excel files
    entry_log_data = read_excel("DataBase.xlsx", "Entry Log", entry_log_columns)
    sales_data = read_excel("Sales.xlsx", "Sales", sales_columns)
    inventory_data = read_excel("Inventory.xlsx", "Inventory", inventory_columns)

    # Filter and sort Entry Log data
    filtered_entry_log_data = []
    for item in entry_log_data:
        # Check if the search query is present in any of the columns
        if any(search_query in str(value).lower() for value in item.values()):
            filtered_entry_log_data.append(item)

    if current_tab == 0:
        if sort_column and sort_order:
            if sort_column == " ":
                sort_column = "ID"
            reverse_order = False if sort_order == "Ascending" else True
            if "Price" in sort_column:
                # Sort by numerical value without "RM" prefix
                filtered_entry_log_data.sort(key=lambda x: float(x[sort_column][2:]), reverse=reverse_order)
            else:
                # Sort by other columns as usual
                filtered_entry_log_data.sort(key=lambda x: x[sort_column], reverse=reverse_order)



    # Update Entry Log tab
    entry_log_tree.delete(*entry_log_tree.get_children())
    for item in filtered_entry_log_data:
        entry_log_tree.insert("", "end", values=list(item.values()))

    # Filter and sort Sales data
    filtered_sales_data = []
    for item in sales_data:
        if any(search_query in str(value).lower() for value in item.values()):
            filtered_sales_data.append(item)

    if current_tab == 1:
        if sort_column == " ":
            sort_column = "Name"
        if sort_column and sort_order:
            reverse_order = False if sort_order == "Ascending" else True
            if "Price" in sort_column:
                # Sort by numerical value without "RM" prefix
                filtered_sales_data.sort(key=lambda x: float(x[sort_column][2:]), reverse=reverse_order)
            else:
                # Sort by other columns as usual
                filtered_sales_data.sort(key=lambda x: x[sort_column], reverse=reverse_order)

    # Update Sales tab
    sales_tree.delete(*sales_tree.get_children())
    for item in filtered_sales_data:
        sales_tree.insert("", "end", values=list(item.values()))

    # Filter and sort Inventory data
    filtered_inventory_data = []
    for item in inventory_data:
        if any(search_query in str(value).lower() for value in item.values()):
            filtered_inventory_data.append(item)

    if current_tab == 2:
        if sort_column == " ":
            sort_column = "Name"
        if sort_column and sort_order:
            reverse_order = False if sort_order == "Ascending" else True
            if "Price" in sort_column:
                # Sort by numerical value without "RM" prefix
                filtered_inventory_data.sort(key=lambda x: float(x[sort_column][2:]), reverse=reverse_order)
            else:
                # Sort by other columns as usual
                filtered_inventory_data.sort(key=lambda x: x[sort_column], reverse=reverse_order)

    # Update Inventory tab
    inventory_tree.delete(*inventory_tree.get_children())
    for item in filtered_inventory_data:
        inventory_tree.insert("", "end", values=list(item.values()))



# Function to open the data entry script
def New_Entry_log():
    df = pd.read_excel("DataBase.xlsx")

    # Function to handle the button click and save data to the database
    def save_data():
        # Load the existing Excel file
        excel_file_path = "DataBase.xlsx"
        sheet_name = "Entry Log"

        item = item_entry.get()
        company_brand = company_brand_var.get()
        unit = int(unit_entry.get())
        price = "RM" + str("{:.2f}".format(float(price_entry.get())))
        date = datetime.datetime.now().strftime("%Y-%m-%d") # Retrieve the date from the DateEntry widget
        max_id = df["ID"].max()  # Get the maximum ID value
        new_id = max_id + 1 if not pd.isna(max_id) else 1  # Calculate the new ID

        data = {"ID": new_id, "Item": item, "Company Brand": company_brand, "Unit": unit, "Price": price, "Date": date}

        # Load and handle the Entry Log Excel file
        try:
            df_existing = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            df_existing = pd.DataFrame(columns=["ID", "Item", "Company Brand", "Unit", "Price", "Date"])

        df_to_append = pd.DataFrame([data], columns=["ID", "Item", "Company Brand", "Unit", "Price", "Date"])
        df_updated = pd.concat([df_existing, df_to_append], ignore_index=True)
        df_updated.to_excel(excel_file_path, sheet_name=sheet_name, index=False, engine='openpyxl')

        # Load and handle the Inventory Excel file
        try:
            inventory_df = pd.read_excel(inventory_excel_file_path, sheet_name=inventory_sheet_name)
        except FileNotFoundError:
            inventory_df = pd.DataFrame(columns=["Name", "Unit", "Company", "Buying Price", "Selling Price", "Date"])

        # Check if the item already exists in the inventory
        existing_item_index = inventory_df.index[inventory_df["Name"] == item]
        if not existing_item_index.empty:
            # Update the unit and date in the inventory for the existing item
            existing_item_index = existing_item_index[0]  # Get the first matching index
            existing_unit = inventory_df.at[existing_item_index, "Unit"]
            updated_unit = existing_unit + ", " + unit
            inventory_df.at[existing_item_index, "Unit"] = updated_unit
            inventory_df.at[existing_item_index, "Date"] = date
        else:
            # Create a new entry in the inventory for the item
            new_inventory_data = {
                "Name": item,
                "Unit": unit,
                "Company": company_brand,
                "Buying Price": price,
                "Selling Price": "",  # You can set this as needed
                "Date": date,
            }
            inventory_df = inventory_df.append(new_inventory_data, ignore_index=True)

        # Save the updated inventory data to the Excel file
        inventory_df.to_excel(inventory_excel_file_path, sheet_name=inventory_sheet_name, index=False,
                              engine='openpyxl')

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

# Function to open the data entry script
def New_Sales():
    # Function to handle the button click and save data to the database
    def save_data():
        # Load the existing Excel file
        excel_file_path = "Sales.xlsx"
        sheet_name = "Sales"

        name = name_entry.get()
        unit = int(Unit_var.get())
        company = company_entry.get()
        selling_price = "RM" + str("{:.2f}".format(float(price_entry.get())))
        date = datetime.datetime.now().strftime("%Y-%m-%d")  # Retrieve the date from the DateEntry widget

        # Load and handle the Inventory Excel file
        try:
            inventory_df = pd.read_excel(inventory_excel_file_path, sheet_name=inventory_sheet_name)
        except FileNotFoundError:
            inventory_df = pd.DataFrame(columns=["Name", "Unit", "Company", "Buying Price", "Selling Price", "Date"])

        # Check if the item exists in the inventory
        inventory_item = inventory_df[(inventory_df["Name"] == name) & (inventory_df["Company"] == company)]
        if inventory_item.empty:
            messagebox.showinfo("Error", "Item not found in inventory")
            return

        # Check if there's enough stock in inventory
        inventory_unit = int(inventory_item.iloc[0]["Unit"])
        if unit > inventory_unit:
            messagebox.showinfo("Error", "Not enough stock in inventory")
            return

        # Update the inventory data
        inventory_item_index = inventory_item.index[0]
        updated_inventory_unit = inventory_unit - unit
        inventory_df.at[inventory_item_index, "Unit"] = updated_inventory_unit
        inventory_df.at[inventory_item_index, "Selling Price"] = selling_price
        inventory_df.at[inventory_item_index, "Date"] = date

        # Save the updated inventory data back to the Excel file
        inventory_df.to_excel(inventory_excel_file_path, sheet_name=inventory_sheet_name, index=False,
                              engine='openpyxl')

        # Create the data for the sales entry
        data = {"Name": name, "Unit": unit, "Company": company, "Selling Price": selling_price, "Date": date}

        # Load and handle the Sales Excel file
        try:
            sales_df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            sales_df = pd.DataFrame(columns=["Name", "Unit", "Company", "Selling Price", "Date"])

        # Concatenate the existing and new data for sales
        df_to_append = pd.DataFrame([data], columns=["Name", "Unit", "Company", "Selling Price", "Date"])
        df_updated = pd.concat([sales_df, df_to_append], ignore_index=True)

        # Save the updated sales data to the Excel file
        df_updated.to_excel(excel_file_path, sheet_name=sheet_name, index=False, engine='openpyxl')

        # Display a message box to confirm the data is saved
        update_data()
        messagebox.showinfo("Success", "Data saved successfully!")

    # Create the main window
    root = tk.Tk()
    root.title("Database Input System")

    # Create and arrange widgets
    item_label = ttk.Label(root, text="Name:")
    name_entry = ttk.Entry(root)

    unit_label = ttk.Label(root, text="Unit:")
    Unit_var = ttk.Entry(root)
    # company_brands = companylist
    # company_brand_var = tk.StringVar()
    # company_brand_dropdown = ttk.Combobox(root, textvariable=company_brand_var, values=company_brands)

    company_label = ttk.Label(root, text="Company:")
    company_entry = ttk.Entry(root)

    price_label = ttk.Label(root, text="Selling Price:")
    price_entry = ttk.Entry(root)

    save_button = ttk.Button(root, text="Save", command=save_data)

    # Arrange widgets using grid layout
    item_label.grid(row=0, column=0, padx=10, pady=5)
    name_entry.grid(row=0, column=1, padx=10, pady=5)

    unit_label.grid(row=2, column=0, padx=10, pady=5)
    Unit_var.grid(row=2, column=1, padx=10, pady=5)

    company_label.grid(row=1, column=0, padx=10, pady=5)
    company_entry.grid(row=1, column=1, padx=10, pady=5)

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
                                        values=(" ", "ID", "Item", "Company Brand", "Price", "Date"))
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
    entry_log_tree.column("ID", width=50, anchor="w")
    for col in entry_log_columns:
        entry_log_tree.heading(col, text=col)

    for item in entry_log_data:
        entry_log_tree.insert("", "end", values=list(item.values()))
    entry_log_tree.pack(fill="both", expand=True)
    notebook.add(entry_log_frame, text="Entry Log")

    # Create a New Data Entry button
    new_data_entry_button = ttk.Button(entry_log_frame, text="New Data Entry", command=New_Entry_log)
    new_data_entry_button.pack(pady=10)  # Adjust the padding as needed

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

    new_data_entry_button = ttk.Button(sales_frame, text="New Data Entry", command=New_Sales)
    new_data_entry_button.pack(pady=10)  # Adjust the padding as needed

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

    # Bind the tab change event to the update_sort_column_combobox function
    notebook.bind("<<NotebookTabChanged>>", lambda event: update_sort_column_combobox())

    # Update the initial values of the sort_column_combobox based on the first tab
    update_sort_column_combobox()


    # Start the main loop
    root.mainloop()
