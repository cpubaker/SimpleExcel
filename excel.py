# excel.py
import tkinter as tk
from tkinter import messagebox
from utils import (
    evaluate_and_update_dependents,
    replace_cell_references,
    save_to_json,
    load_from_json,
    cloud_save,
    cloud_load
)

class SimpleExcel:
    ROWS = 5
    COLS = 6

    def __init__(self, root):
        self.root = root
        self.root.title("Simple Excel")

        # Initialize self.data, self.formulas, and self.calculated_values
        self.data = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]
        self.formulas = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]
        self.calculated_values = [['' for _ in range(self.COLS)] for _ in range(self.ROWS)]

        # Variable to store the selected cell coordinates
        self.selected_cell = None
        # Variable to store the last edited cell coordinates
        self.last_edited_cell = None

        # Create Save, Load, Cloud Save, Cloud Load, and Entry widgets in a top frame
        self.create_button_frame()

        self.create_labels_frame()

        # Create a table (editable grid) with labels above it
        self.create_table()

        # Bind the new method to the cell selection event
        self.table_frame.bind('<FocusIn>', self.on_cell_selection)

        # Bind the new method to the focus-out event
        self.root.bind('<FocusOut>', self.on_focus_out)

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

                # Attach an event handler to capture the focus-out event
                self.entries[i][j].bind('<FocusOut>', lambda event, row=i, col=j: self.on_focus_out(event))

    def on_enter_key(self, row, col):
        # Evaluate and update dependent cells if there's a formula
        evaluate_and_update_dependents(self, row, col)
    
        # Set the selected cell coordinates
        self.selected_cell = (row, col)
    
        # Display the content of the selected cell in the entry widget
        self.selected_entry.config(state=tk.NORMAL)  # Enable the entry field
    
        if self.selected_cell == (row, col):
            # Show the formula if there is a formula, otherwise show the value
            if self.formulas[row][col]:
                formula = '=' + self.formulas[row][col]
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, formula)
    
                # Update the cell to show the formula result
                self.entries[row][col].delete(0, tk.END)
                self.entries[row][col].insert(0, str(self.calculated_values[row][col]))
            else:
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, self.data[row][col])

        # Update the flag to indicate that the cell is being edited
        self.editing_cell = True

        # Update the last edited cell
        self.last_edited_cell = (row, col)

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

        if self.selected_cell == (row, col):
            # Show the formula if there is a formula, otherwise show the value
            if self.formulas[row][col]:
                formula = '=' + self.formulas[row][col]
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, formula)

                # Update the cell to show the formula
                self.entries[row][col].delete(0, tk.END)
                self.entries[row][col].insert(0, formula)

                # If the formula has not changed since the last edit, update the calculated value
                if self.last_edited_cell == (row, col):
                    evaluate_and_update_dependents(self, row, col)
            else:
                self.selected_entry.delete(0, tk.END)
                self.selected_entry.insert(0, self.data[row][col])

        # Update the flag to indicate that the cell is being edited
        self.editing_cell = True

        # Update the last edited cell
        self.last_edited_cell = (row, col)

    def on_selected_entry_change(self, event):
        # Update the data when the separate entry field value changes
        if self.selected_cell:
            new_value = self.selected_entry.get()
            self.data[self.selected_cell[0]][self.selected_cell[1]] = new_value
            # Update the corresponding grid cell if it exists
            self.entries[self.selected_cell[0]][self.selected_cell[1]].delete(0, tk.END)
            self.entries[self.selected_cell[0]][self.selected_cell[1]].insert(0, new_value)

    # Add a new method to handle cell selection without pressing Enter
    def on_cell_selection(self, event):
        # Set the selected cell coordinates
        row, col = self.selected_cell

        # Display the content of the selected cell in the entry widget
        self.selected_entry.config(state=tk.NORMAL)  # Enable the entry field

    def on_focus_out(self, event):
        # Process the formula for the last edited cell when an entry widget loses focus
        if self.last_edited_cell:
            row, col = self.get_selected_cell()
            if self.formulas[row][col]:
                evaluate_and_update_dependents(self, row, col)

        # Reset the last edited cell when an entry widget loses focus
        self.last_edited_cell = None

    # Add a new method to get the selected cell coordinates
    def get_selected_cell(self):
        return self.selected_cell    

    def save_to_json(self):
        save_to_json(self)

    def load_from_json(self):
        load_from_json(self)

    def cloud_save(self):
        cloud_save(self)

    def cloud_load(self):
        cloud_load(self)
