import tkinter as tk
from tkinter import messagebox
from utils import evaluate_and_update_dependents, replace_cell_references, save_to_json, load_from_json, cloud_save, cloud_load
from mongodb_utils import cloud_save_to_mongodb, cloud_load_from_mongodb

class SimpleExcel:
    ROWS = 5
    COLS = 6

    def __init__(self, root):
        self.root = root
        self.root.title("Simple Excel")

        self.data = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]
        self.formulas = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]

        self.selected_cell = None
        self.last_edited_cell = None

        self.create_button_frame()
        self.create_labels_frame()
        self.create_table()

        self.table_frame.bind('<FocusIn>', self.handle_cell_selection)
        self.root.bind('<FocusOut>', self.handle_focus_out)

    def create_button_frame(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(side=tk.TOP)

        save_button = tk.Button(button_frame, text="Save", command=self.save_to_json)
        save_button.pack(side=tk.LEFT, padx=5, pady=5)

        load_button = tk.Button(button_frame, text="Load", command=self.load_from_json)
        load_button.pack(side=tk.LEFT, padx=(0, 5), pady=5)

        cloud_save_button = tk.Button(button_frame, text="Cloud Save", command=self.cloud_save)
        cloud_save_button.pack(side=tk.LEFT, padx=5, pady=5)

        cloud_load_button = tk.Button(button_frame, text="Cloud Load", command=self.cloud_load)
        cloud_load_button.pack(side=tk.LEFT, padx=(0, 5), pady=5)

        self.selected_entry = tk.Entry(button_frame, width=10, state=tk.DISABLED)
        self.selected_entry.pack(side=tk.LEFT, padx=(10, 5), pady=5)
        self.selected_entry.bind('<KeyRelease>', self.handle_selected_entry_change)

    def create_labels_frame(self):
        labels_frame = tk.Frame(self.root)
        labels_frame.pack(side=tk.TOP)

        for col, label_text in enumerate(['A', 'B', 'C', 'D', 'E']):
            label = tk.Label(labels_frame, text=label_text)
            label.pack(side=tk.LEFT, padx=26)

    def create_table(self):
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(side=tk.TOP)

        self.entries = [[tk.Entry(self.table_frame, width=10) for _ in range(self.COLS)] for _ in range(self.ROWS)]

        for i in range(self.ROWS):
            row_label = tk.Label(self.table_frame, text=str(i + 1), width=2, anchor=tk.W, justify=tk.RIGHT)
            row_label.grid(row=i + 1, column=0)

            for j in range(self.COLS):
                entry = self.entries[i][j]
                entry.grid(row=i + 1, column=j + 1)

                entry.bind('<KeyRelease>', lambda event, row=i, col=j: self.handle_cell_change(row, col))
                entry.bind('<Button-1>', lambda event, row=i, col=j: self.handle_cell_click(row, col))
                entry.bind('<Return>', lambda event, row=i, col=j: self.handle_enter_key(row, col))
                entry.bind('<FocusOut>', lambda event, row=i, col=j: self.handle_focus_out(event))

    def handle_enter_key(self, row, col):
        self.selected_cell = (row, col)
        self.selected_entry.config(state=tk.NORMAL)
    
        if self.selected_cell == (row, col):
            formula = self.selected_entry.get()
    
            # Remove any existing '=' sign
            formula = formula.lstrip('=')
    
            # Store the formula in the formulas list
            self.formulas[row][col] = formula
    
            # If there is a formula, display it in the entry field with '=' sign
            if formula:
                # Add '=' sign if not already present
                formula_with_equals = '=' + formula if not formula.startswith('=') else formula
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, formula_with_equals)
    
        self.editing_cell = True
        self.last_edited_cell = (row, col)
        self.handle_cell_change(row, col)  # Added to trigger recalculation


    def handle_cell_change(self, row, col):
        old_value = self.data[row][col]
        new_value = self.entries[row][col].get()

        # Check if the value has changed
        if old_value != new_value:
            # If the cell contains a formula, evaluate and update
            if self.formulas[row][col]:
                evaluate_and_update_dependents(self, row, col)
            else:
                self.data[row][col] = new_value

            # Update the separate entry field if the selected cell matches the changed cell
            if self.selected_cell == (row, col):
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, new_value)

            # Recalculate dependents for all cells with formulas
            for i in range(self.ROWS):
                for j in range(self.COLS):
                    if self.formulas[i][j]:
                        evaluate_and_update_dependents(self, i, j)


    def handle_cell_click(self, row, col):
        self.selected_cell = (row, col)
        self.selected_entry.config(state=tk.NORMAL)

        if self.selected_cell == (row, col):
            if self.formulas[row][col]:
                formula = '=' + self.formulas[row][col]
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, formula)

                self.entries[row][col].delete(0, tk.END)
                self.entries[row][col].insert(0, formula)

                if self.last_edited_cell == (row, col):
                    evaluate_and_update_dependents(self, row, col)
            else:
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, self.data[row][col])

        self.editing_cell = True
        self.last_edited_cell = (row, col)

    def handle_selected_entry_change(self, event):
        if self.selected_cell:
            new_value = self.selected_entry.get()
            self.data[self.selected_cell[0]][self.selected_cell[1]] = new_value
            self.entries[self.selected_cell[0]][self.selected_cell[1]].delete(0, tk.END)
            self.entries[self.selected_cell[0]][self.selected_cell[1]].insert(0, new_value)

    def handle_cell_selection(self, event):
        row, col = self.selected_cell
        self.selected_entry.config(state=tk.NORMAL)

    def handle_focus_out(self, event):
        if self.last_edited_cell:
            row, col = self.get_selected_cell()
            if self.formulas[row][col]:
                evaluate_and_update_dependents(self, row, col)

        self.last_edited_cell = None

    def get_selected_cell(self):
        return self.selected_cell

    def save_to_json(self):
        save_to_json(self)

    def load_from_json(self):
        load_from_json(self)

    def cloud_save(self):
        success, message = cloud_save_to_mongodb(self.data, self.formulas, self.calculated_values)
        if success:
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)

    def cloud_load(self):
        success, result = cloud_load_from_mongodb()
        if success:
            loaded_data = result
            self.data = loaded_data.get("data", [['' for _ in range(self.COLS)] for _ in range(self.ROWS)])
            self.formulas = loaded_data.get("formulas", [['' for _ in range(self.COLS)] for _ in range(self.ROWS)])
            self.calculated_values = loaded_data.get("calculated_values", [['' for _ in range(self.COLS)] for _ in range(self.ROWS)])

            # Update the GUI with the loaded data
            for i in range(self.ROWS):
                for j in range(self.COLS):
                    self.entries[i][j].delete(0, tk.END)
                    self.entries[i][j].insert(0, str(self.data[i][j]))

            messagebox.showinfo("Success", "Data loaded from MongoDB.")
        else:
            messagebox.showerror("Error", result)

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleExcel(root)
    root.mainloop()
