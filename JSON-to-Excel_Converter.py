import json
import csv
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, List, Any, Set
import os

class JSONConverter:
    def __init__(self):
        self.json_data = None
        self.file_path = None
        self.all_keys = set()
        self.nested_keys = {}
        
    def select_json_file(self) -> bool:
        """Prompt user to select a JSON file."""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        file_path = filedialog.askopenfilename(
            title="Select JSON file",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            print("No file selected.")
            return False
            
        self.file_path = Path(file_path)
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                self.json_data = json.load(file)
            print(f"Successfully loaded: {self.file_path.name}")
            return True
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON file: {e}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {e}")
            return False
    
    def extract_all_keys(self, data: Any, parent_key: str = "") -> None:
        """Recursively extract all keys and subkeys from JSON data."""
        if isinstance(data, dict):
            for key, value in data.items():
                full_key = f"{parent_key}.{key}" if parent_key else key
                self.all_keys.add(full_key)
                
                if isinstance(value, (dict, list)):
                    self.extract_all_keys(value, full_key)
                    
        elif isinstance(data, list) and data:
            # Check all items in the list to capture all possible keys
            for item in data:
                if isinstance(item, dict):
                    self.extract_all_keys(item, parent_key)
                elif isinstance(item, list):
                    self.extract_all_keys(item, parent_key)
    
    def get_leaf_keys_only(self) -> List[str]:
        """Filter keys to show only leaf keys (keys without subkeys)."""
        leaf_keys = []
        sorted_keys = sorted(list(self.all_keys))
        
        for key in sorted_keys:
            # Check if this key has any subkeys
            has_subkeys = any(other_key.startswith(key + ".") for other_key in sorted_keys if other_key != key)
            
            if not has_subkeys:
                leaf_keys.append(key)
        
        return leaf_keys
    
    def display_available_keys(self) -> None:
        """Display all available keys and subkeys."""
        if isinstance(self.json_data, list):
            print("JSON Structure: Array of objects")
            self.extract_all_keys(self.json_data)
        elif isinstance(self.json_data, dict):
            print("JSON Structure: Single object")
            self.extract_all_keys(self.json_data)
        else:
            print("JSON Structure: Simple value")
            return
        
        # Get only leaf keys (those without subkeys)
        leaf_keys = self.get_leaf_keys_only()
        
        print("\nAvailable keys for export:")
        print("-" * 40)
        for i, key in enumerate(leaf_keys, 1):
            print(f"{i:2d}. {key}")
        
    def get_user_selection(self) -> List[str]:
        """Get user's selection of keys to export."""
        leaf_keys = self.get_leaf_keys_only()
        
        print(f"\nSelect keys to export (1-{len(leaf_keys)}):")
        print("Enter numbers separated by commas (e.g., 1,3,5-7) or 'all' for all keys:")
        
        while True:
            selection = input("Your selection: ").strip()
            
            if selection.lower() == 'all':
                return leaf_keys
            
            try:
                selected_indices = []
                parts = selection.split(',')
                
                for part in parts:
                    part = part.strip()
                    if '-' in part:
                        # Handle range like "5-7"
                        start, end = map(int, part.split('-'))
                        selected_indices.extend(range(start, end + 1))
                    else:
                        selected_indices.append(int(part))
                
                # Convert to 0-based indices and get corresponding keys
                selected_keys = []
                for idx in selected_indices:
                    if 1 <= idx <= len(leaf_keys):
                        selected_keys.append(leaf_keys[idx - 1])
                    else:
                        print(f"Invalid index: {idx}")
                        raise ValueError()
                
                return selected_keys
                
            except ValueError:
                print("Invalid selection. Please try again.")
    
    def get_nested_value(self, data: Any, key_path: str) -> Any:
        """Get value from nested dictionary using dot notation."""
        keys = key_path.split('.')
        current = data
        
        try:
            for key in keys:
                if isinstance(current, dict):
                    current = current.get(key)
                elif isinstance(current, list):
                    # Handle arrays - we'll return the list itself or extract from all items
                    if key.isdigit():
                        # If key is numeric, treat as array index
                        idx = int(key)
                        current = current[idx] if 0 <= idx < len(current) else None
                    else:
                        # Extract key from all items in the array
                        extracted_values = []
                        for item in current:
                            if isinstance(item, dict) and key in item:
                                extracted_values.append(item[key])
                        current = extracted_values if extracted_values else None
                else:
                    current = None
                    
                if current is None:
                    break
                    
            return current
        except (KeyError, IndexError, AttributeError, ValueError):
            return None
    
    def flatten_data_for_export(self, selected_keys: List[str]) -> List[Dict[str, Any]]:
        """Flatten JSON data for CSV/XLSX export."""
        flattened_data = []
        
        if isinstance(self.json_data, list):
            # Array of objects - process each item
            for item in self.json_data:
                row = {}
                for key in selected_keys:
                    value = self.get_nested_value(item, key)
                    # Convert complex objects to string representation
                    if isinstance(value, (dict, list)):
                        if isinstance(value, list) and len(value) == 1 and not isinstance(value[0], (dict, list)):
                            # If it's a single-item list with a simple value, extract the value
                            value = value[0]
                        else:
                            value = json.dumps(value, ensure_ascii=False)
                    row[key] = value
                flattened_data.append(row)
                
        elif isinstance(self.json_data, dict):
            # Single object - but it might contain arrays that need to be expanded
            # Check if any selected keys contain arrays that should create multiple rows
            array_keys = []
            max_length = 1
            
            for key in selected_keys:
                value = self.get_nested_value(self.json_data, key)
                if isinstance(value, list) and value and not isinstance(value[0], (dict, list)):
                    array_keys.append(key)
                    max_length = max(max_length, len(value))
            
            # Create rows based on the longest array
            for i in range(max_length):
                row = {}
                for key in selected_keys:
                    value = self.get_nested_value(self.json_data, key)
                    
                    if key in array_keys and isinstance(value, list):
                        # Get the i-th element if it exists
                        row[key] = value[i] if i < len(value) else None
                    else:
                        # For non-array values, repeat the same value for each row
                        if isinstance(value, (dict, list)) and key not in array_keys:
                            value = json.dumps(value, ensure_ascii=False)
                        row[key] = value
                        
                flattened_data.append(row)
            
        return flattened_data
    
    def export_to_csv(self, data: List[Dict[str, Any]], selected_keys: List[str]) -> None:
        """Export data to CSV file."""
        output_file = self.file_path.with_suffix('.csv')
        
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=selected_keys)
                writer.writeheader()
                writer.writerows(data)
            
            print(f"CSV file created: {output_file}")
            
        except Exception as e:
            print(f"Error creating CSV file: {e}")
    
    def export_to_xlsx(self, data: List[Dict[str, Any]], selected_keys: List[str]) -> None:
        """Export data to XLSX file."""
        output_file = self.file_path.with_suffix('.xlsx')
        
        try:
            df = pd.DataFrame(data, columns=selected_keys)
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"XLSX file created: {output_file}")
            
        except Exception as e:
            print(f"Error creating XLSX file: {e}")
            print("Note: Make sure you have openpyxl installed: pip install openpyxl")
    
    def run(self) -> None:
        """Main execution method."""
        print("JSON to CSV/XLSX Converter")
        print("=" * 30)
        
        # Step 1: Select JSON file
        if not self.select_json_file():
            return
        
        # Step 2: Display available keys
        self.display_available_keys()
        
        leaf_keys = self.get_leaf_keys_only()
        if not leaf_keys:
            print("No exportable keys found in the JSON file.")
            return
        
        # Step 3: Get user selection
        selected_keys = self.get_user_selection()
        
        if not selected_keys:
            print("No keys selected for export.")
            return
        
        print(f"\nSelected keys: {', '.join(selected_keys)}")
        
        # Step 4: Flatten data and export
        flattened_data = self.flatten_data_for_export(selected_keys)
        
        if not flattened_data:
            print("No data to export.")
            return
        
        # Export to both CSV and XLSX
        self.export_to_csv(flattened_data, selected_keys)
        self.export_to_xlsx(flattened_data, selected_keys)
        
        print(f"\nExport completed! Generated {len(flattened_data)} rows.")

def main():
    """Main function to run the converter."""
    converter = JSONConverter()
    converter.run()

if __name__ == "__main__":
    # Required dependencies
    required_packages = ['pandas', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("Missing required packages. Please install them using:")
        print(f"pip install {' '.join(missing_packages)}")
        print("\nThen run the script again.")
    else:
        main()