import openpyxl
import os

class Add_NewData():
    def __init__(self, file_name="./storage/Librarie_Saint_Laurent.xlsx", sheet_name="book_data"):
        self.file_name = file_name
        self.workbook = None
        self.worksheet = None
        self.sheet_name = sheet_name

    def Add_book(self, title="Sample Book", author="Unknown", publication_year=2025):
        if os.path.exists(self.file_name):
            # Load the existing workbook
            self.workbook = openpyxl.load_workbook(self.file_name)
            self.worksheet = self.workbook.active  
        else:
            print(f"File '{self.file_name}' not found. Creating a new file.")
            # Create a new workbook and add headers
            self.workbook = openpyxl.Workbook()
            self.worksheet = self.workbook.active
            self.worksheet.append(["Title", "Author", "Publication Year"])  # Column headers

        # Append the new book data
        self.worksheet.append([title, author, publication_year])
        
        # Save the workbook
        self.workbook.save(self.file_name)
        print(f"Book '{title}' added successfully to {self.file_name}!")

def main():
    manager = Add_NewData()
    manager.Add_book()

if __name__ == "__main__":
    main()
