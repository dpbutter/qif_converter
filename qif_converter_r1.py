import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
from datetime import datetime
from ttkbootstrap import Style

class QIFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to QIF Converter")
        self.csv_filename = None

        # Set the style to use ttkbootstrap
        self.style = Style(theme='flatly')

        # Create a button to select the CSV file using ttk.Button
        self.select_file_btn = ttk.Button(root, text="Select CSV File", command=self.select_csv_file, bootstyle="success")
        self.select_file_btn.pack(pady=10)

        # Placeholder for the frame that will hold the mappings
        self.mapping_frame = None

        # Initialize the widgets for QIF type and amount interpretation but don't display them yet
        self.qif_type_var = tk.StringVar(value="CCard")
        self.amount_sign_var = tk.StringVar(value="Positive = withdrawal")

        # Button to save QIF file
        self.save_qif_btn = ttk.Button(root, text="Save QIF File", command=self.save_qif_file, state=tk.DISABLED, bootstyle="primary")

        # Define QIF options
        self.qif_options = ["Date", "Payee", "Amount", "Category", "Not Used"]
        self.mapping_vars = []  # Store dropdown variables
        self.header_present = True  # Track if the CSV file has a header

    def select_csv_file(self):
        self.csv_filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not self.csv_filename:
            return
        
        # Ask user if the CSV has a header
        self.header_present = messagebox.askyesno("Header Present?", "Does the CSV file have a header row?")
        self.load_csv_fields(self.header_present)

    def load_csv_fields(self, header_present):
        # Clear the previous frame if it exists
        if self.mapping_frame is not None:
            self.mapping_frame.destroy()

        # Create a new frame for mappings
        self.mapping_frame = tk.Frame(self.root)
        self.mapping_frame.pack(pady=10)

        # Read the CSV and collect the top 5 rows for preview
        with open(self.csv_filename, 'r') as csv_file:
            reader = csv.reader(csv_file)
            rows = list(reader)
            
            if header_present:
                headers = rows[0]
                top_entries = rows[1:6]  # Get the next 5 rows for preview
            else:
                headers = [f"Column {i+1}" for i in range(len(rows[0]))]
                top_entries = rows[:5]  # Get the first 5 rows for preview

        # Display QIF Header Type and Amount Interpretation labels and dropdowns
        qif_type_label = tk.Label(self.mapping_frame, text="QIF Header Type:")
        qif_type_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        qif_type_dropdown = ttk.Combobox(self.mapping_frame, textvariable=self.qif_type_var, values=["Bank", "Cash", "CCard"], state="readonly")
        qif_type_dropdown.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # Display headers and create dropdowns for mapping
        self.mapping_vars.clear()  # Reset mapping_vars
        for index, header in enumerate(headers):
            label = tk.Label(self.mapping_frame, text=header)
            label.grid(row=0, column=index + 1, padx=5, pady=5, sticky="w")

            # Set default option to "Not Used"
            var = tk.StringVar(value="Not Used")
            dropdown = ttk.Combobox(self.mapping_frame, textvariable=var, values=self.qif_options, state="readonly")
            dropdown.grid(row=1, column=index + 1, padx=5, pady=5, sticky="ew")
            dropdown.bind("<<ComboboxSelected>>", self.check_unique_selection)
            self.mapping_vars.append((header, var, dropdown))  # Store dropdown reference

        amount_sign_label = tk.Label(self.mapping_frame, text="Amount Interpretation:")
        amount_sign_label.grid(row=0, column=len(headers)+1, padx=5, pady=5, sticky="w")
        amount_sign_dropdown = ttk.Combobox(self.mapping_frame, textvariable=self.amount_sign_var, values=["Positive = withdrawal", "Positive = deposit"], state="readonly")
        amount_sign_dropdown.grid(row=1, column=len(headers)+1, padx=5, pady=5, sticky="ew")

        # Enable the "Save QIF File" button
        self.save_qif_btn.config(state=tk.NORMAL)

        # Re-pack the "Save QIF File" button to ensure it is always below the mapping frame
        self.save_qif_btn.pack_forget()  # Unpack it first to reposition
        self.save_qif_btn.pack(pady=20)

        # Display the top 5 entries below the dropdowns, aligned with headers
        for row_index, entry in enumerate(top_entries):
            for col_index, value in enumerate(entry):
                # Truncate the value to a maximum of 20 characters for display
                truncated_value = value[:20] + '...' if len(value) > 20 else value
                entry_label = tk.Label(self.mapping_frame, text=truncated_value, borderwidth=1, relief="solid")
                entry_label.grid(row=row_index + 2, column=col_index + 1, padx=5, pady=5, sticky="ew")

    def check_unique_selection(self, event):
        selected_value = event.widget.get()  # Get the selected value
        current_dropdowns = [var.get() for _, var, _ in self.mapping_vars]  # Get current selections

        # Check for duplicates, skip if the selected value is "Not Used"
        if selected_value != "Not Used" and current_dropdowns.count(selected_value) > 1:
            messagebox.showwarning("Duplicate Selection", f"The option '{selected_value}' has already been selected.")
            # Revert to the last valid selection
            last_selection = self.mapping_vars[current_dropdowns.index(selected_value)][1].get()
            event.widget.set(last_selection)

    def save_qif_file(self):
        # Get the filename for the output QIF file
        qif_filename = filedialog.asksaveasfilename(defaultextension=".qif", filetypes=[("QIF Files", "*.qif")])
        if not qif_filename:
            return

        # Read the selected field mappings
        mappings = {header: var.get() for header, var, _ in self.mapping_vars}
        qif_type = self.qif_type_var.get()
        amount_sign_option = self.amount_sign_var.get()

        # Convert the CSV file to QIF using the selected mappings
        try:
            with open(self.csv_filename, 'r') as csv_file:
                if self.header_present:
                    csv_reader = csv.DictReader(csv_file)
                else:
                    csv_reader = csv.reader(csv_file)
                
                with open(qif_filename, 'w') as qif_file:
                    # Write the QIF header based on user selection
                    qif_file.write(f'!Type:{qif_type}\n')
                    
                    # Process each row in the CSV file
                    for row in csv_reader:
                        # Initialize a dictionary to store QIF fields
                        qif_data = {'Date': None, 'Payee': None, 'Amount': None, 'Category': None}

                        # Map CSV fields to QIF fields based on user selection
                        for index, (csv_field, qif_field) in enumerate(mappings.items()):
                            if qif_field == "Not Used":
                                continue
                            # Handle mapping for non-header CSV files
                            value = row[index] if not self.header_present else row[csv_field]
                            if qif_field == "Date":
                                # Convert date to QIF format (MM/DD/YYYY)
                                qif_data['Date'] = datetime.strptime(value, '%m/%d/%Y').strftime('%m/%d/%Y')
                            elif qif_field == "Payee":
                                qif_data['Payee'] = value
                            elif qif_field == "Amount":
                                # Parse the amount and adjust the sign based on user option
                                amount = float(value)
                                if amount_sign_option == "Positive = withdrawal":
                                    amount = -amount
                                qif_data['Amount'] = f"{amount:.2f}"
                            elif qif_field == "Category":
                                qif_data['Category'] = value

                        # Write the QIF transaction if essential fields are present
                        if qif_data['Date'] and qif_data['Amount']:
                            qif_file.write(f"D{qif_data['Date']}\n")
                            qif_file.write(f"T{qif_data['Amount']}\n")
                            if qif_data['Payee']:
                                qif_file.write(f"P{qif_data['Payee']}\n")
                            if qif_data['Category']:
                                qif_file.write(f"L{qif_data['Category']}\n")
                            qif_file.write("^\n")  # End of transaction

            messagebox.showinfo("Success", "QIF file has been created successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the QIF file: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QIFConverterApp(root)
    root.mainloop()

