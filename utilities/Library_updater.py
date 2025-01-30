from openpyxl import load_workbook
import os
import openpyxl

class UpdateSheet:
    def __init__(self, file_name="./storage/Librarie_Saint_Laurent.xlsx"):
        self.file_name = file_name
        self.workbook = None
        self.worksheet = None
        
    def DatainCells(self):
        if os.path.exists(self.file_name):
                # Load the workbook
            self.workbook = load_workbook(self.file_name)
            self.worksheet = self.workbook.active
                
                # Update cell value
            self.worksheet["B3"] = "Unkown Author"
            
                
                # Save changes using the instance variable
            self.workbook.save(self.file_name)
            print(f"Cell updated successfully in '{self.file_name}'")
        else:
            print(f"Error: File '{self.file_name}' not found")
                