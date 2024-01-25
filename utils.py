# utils.py
import json
import re
import tkinter as tk
from sympy import sympify
from tkinter import messagebox

def evaluate_and_update_dependents(self, row, col):
    # Check if the cell contains a formula
    print("evaluate_and_update_dependents")

    if self.formulas[row][col]:
        formula = self.formulas[row][col]
        try:
            if not formula:
                # Skip evaluation if the formula is empty
                return

            print(f"Evaluating formula: {formula} at cell ({row}, {col})")

            # Replace cell references with their values
            processed_formula = replace_cell_references(formula, self)

            print(f"Processed formula: {processed_formula}")

            # Evaluate the formula using sympy
            result = sympify(processed_formula)

            print(f"Result: {result}")

            # Update the cell value, formula, and calculated value with the result
            self.entries[row][col].delete(0, tk.END)
            self.entries[row][col].insert(0, str(result))
            self.data[row][col] = str(result)
            self.formulas[row][col] = formula
            self.calculated_values[row][col] = str(result)

            # Update other cells that depend on this cell
            for i in range(self.ROWS):
                for j in range(self.COLS):
                    if i != row and j != col and self.formulas[i][j]:
                        print(f"Recursively calling evaluate_and_update_dependents for cell ({i}, {j})")
                        self.evaluate_and_update_dependents(i, j)

        except Exception as e:
            messagebox.showerror("Error", f"Error evaluating formula: {e}")

def replace_cell_references(formula, excel_instance):
    # Replace cell references (e.g., =A1) with their corresponding values
    cell_pattern = re.compile(r'[A-E]\d')
    matches = cell_pattern.findall(formula)
    for match in matches:
        row = int(match[1]) - 1
        col = ord(match[0]) - ord('A')
        cell_value = excel_instance.data[row][col]
        print(f"Replacing {match} with {cell_value} in formula: {formula}")
        formula = formula.replace(match, cell_value)
    return formula


def save_to_json(self):
    # Save data, formulas, and calculated values to a JSON file
    data_to_save = {
        'data': self.data,
        'formulas': self.formulas,
        'calculated_values': self.calculated_values
    }
    with open('excel_data.json', 'w') as json_file:
        json.dump(data_to_save, json_file)
def load_from_json(self):
    try:
        with open('excel_data.json', 'r') as json_file:
            loaded_data = json.load(json_file)
        self.data = loaded_data['data']
        self.formulas = loaded_data['formulas']
        self.calculated_values = loaded_data['calculated_values']
        for i in range(self.ROWS):
            for j in range(self.COLS):
                self.entries[i][j].delete(0, tk.END) 
                self.entries[i][j].insert(0, self.data[i][j])
    except (FileNotFoundError, json.decoder.JSONDecodeError) as e:
        messagebox.showerror("Error", f"Error loading data: {e}")

def cloud_save(self):
    # Placeholder for cloud save functionality
    print("Cloud Save button pressed")
    
def cloud_load(self):
    # Placeholder for cloud load functionality
    print("Cloud Load button pressed")