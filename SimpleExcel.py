import tkinter as tk
from tkinter import messagebox
import json
import re
from sympy import symbols, sympify

class SimpleExcel:
    ROWS = 5
    COLS = 6

    def __init__(self, root):
        self.root = root
        self.root.title("Simple Excel")

        # Initialize self.data
        self.data = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]

        # Variable to store the selected cell coordinates
        self.selected_cell = None

        # Create Save, Load, Cloud Save, Cloud Load, and Entry widgets in a top frame
        self.create_button_frame()

        self.create_labels_frame()

        # Create a table (editable grid) with labels above it
        self.create_table()

    def create_button_frame(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(side=tk.TOP)

        # Save button
        save_button = tk.Button(button_frame, text="Save", command=self.save_to_json)
        save_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Load button
        load_button = tk.Button(button_frame, text="Load", command=self.load_from_json)
        load_button.pack(side=tk.LEFT, padx=(0, 5), pady=5)

        # Cloud Save button
        cloud_save_button = tk.Button(button_frame, text="Cloud Save", command=self.cloud_save)
        cloud_save_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Cloud Load button
        cloud_load_button = tk.Button(button_frame, text="Cloud Load", command=self.cloud_load)
        cloud_load_button.pack(side=tk.LEFT, padx=(0, 5), pady=5)

        # Entry widget for selected cell
        self.selected_entry = tk.Entry(button_frame, width=10, state=tk.DISABLED)
        self.selected_entry.pack(side=tk.LEFT, padx=(10, 5), pady=5)
        # Attach an event handler to capture changes in the separate entry field
        self.selected_entry.bind('<KeyRelease>', self.on_selected_entry_change)

    def create_labels_frame(self):
        labels_frame = tk.Frame(self.root)
        labels_frame.pack(side=tk.TOP)

        for col, label_text in enumerate(['A', 'B', 'C', 'D', 'E']):
            label = tk.Label(labels_frame, text=label_text)
            label.pack(side=tk.LEFT, padx=26)

    def create_table(self):
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(side=tk.TOP)

        # Create entry widgets in each cell
        self.entries = [[tk.Entry(self.table_frame, width=10) for _ in range(self.COLS)] for _ in range(self.ROWS)]

        # Place the entry widgets in the grid
        for i in range(self.ROWS):
            # Add labels for row numbers
            row_label = tk.Label(self.table_frame, text=str(i + 1), width=2, anchor=tk.W, justify=tk.RIGHT)
            row_label.grid(row=i + 1, column=0)

            for j in range(self.COLS):
                self.entries[i][j].grid(row=i + 1, column=j + 1)
                # Attach an event handler to capture changes in the entry
                self.entries[i][j].bind('<KeyRelease>', lambda event, row=i, col=j: self.on_cell_change(row, col))

                # Attach an event handler to capture clicks on the entry
                self.entries[i][j].bind('<Button-1>', lambda event, row=i, col=j: self.on_cell_click(row, col))

                # Attach an event handler to capture the Enter key press
                self.entries[i][j].bind('<Return>', lambda event, row=i, col=j: self.on_enter_key(row, col))

    def on_enter_key(self, row, col):
        # Evaluate and update dependent cells if there's a formula
        self.evaluate_and_update_dependents(row, col)

    def on_cell_change(self, row, col):
        # Update the data when a cell value changes
        self.data[row][col] = self.entries[row][col].get()

        # Update the separate entry field if the selected cell matches the changed cell
        if self.selected_cell == (row, col):
            self.selected_entry.delete(0, tk.END)
            self.selected_entry.insert(0, self.data[row][col])

    def on_cell_click(self, row, col):
        # Set the selected cell coordinates
        self.selected_cell = (row, col)

        # Display the content of the selected cell in the entry widget
        self.selected_entry.config(state=tk.NORMAL)  # Enable the entry field
        self.selected_entry.delete(0, tk.END)
        self.selected_entry.insert(0, self.data[row][col])

    def on_selected_entry_change(self, event):
        # Update the data when the separate entry field value changes
        if self.selected_cell:
            new_value = self.selected_entry.get()
            self.data[self.selected_cell[0]][self.selected_cell[1]] = new_value
            # Update the corresponding grid cell if it exists
            self.entries[self.selected_cell[0]][self.selected_cell[1]].delete(0, tk.END)
            self.entries[self.selected_cell[0]][self.selected_cell[1]].insert(0, new_value)

    def evaluate_and_update_dependents(self, row, col):
        # Check if the cell contains a formula
        if self.data[row][col] and self.data[row][col][0] == '=':
            formula = self.data[row][col][1:]
            try:
                # Replace cell references with their values
                processed_formula = self.replace_cell_references(formula)

                # Evaluate the formula using sympy
                result = sympify(processed_formula)
                
                # Update the cell value with the result
                self.entries[row][col].delete(0, tk.END)
                self.entries[row][col].insert(0, str(result))
                self.data[row][col] = str(result)
                
                # Update other cells that depend on this cell
                for i in range(self.ROWS):
                    for j in range(self.COLS):
                        if i != row and j != col and self.data[i][j] and self.data[i][j][0] == '=':
                            self.evaluate_and_update_dependents(i, j)

            except Exception as e:
                messagebox.showerror("Error", f"Error evaluating formula: {e}")

    def replace_cell_references(self, formula):
        # Replace cell references (e.g., =A1) with their corresponding values
        cell_pattern = re.compile(r'[A-E]\d')
        matches = cell_pattern.findall(formula)

        for match in matches:
            row = int(match[1]) - 1
            col = ord(match[0]) - ord('A')
            cell_value = self.data[row][col]
            formula = formula.replace(match, cell_value)

        return formula

    def save_to_json(self):
        # Save data to a JSON file
        with open('excel_data.json', 'w') as json_file:
            json.dump(self.data, json_file)

    def load_from_json(self):
        try:
            with open('excel_data.json', 'r') as json_file:
                loaded_data = json.load(json_file)

            for i in range(self.ROWS):
                for j in range(self.COLS):
                    self.entries[i][j].delete(0, tk.END)
                    self.entries[i][j].insert(0, loaded_data[i][j])
                    self.data[i][j] = loaded_data[i][j]

        except (FileNotFoundError, json.decoder.JSONDecodeError) as e:
            messagebox.showerror("Error", f"Error loading data: {e}")

    def cloud_save(self):
        # Placeholder for cloud save functionality
        print("Cloud Save button pressed")

    def cloud_load(self):
        # Placeholder for cloud load functionality
        print("Cloud Load button pressed")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleExcel(root)
    root.mainloop()
