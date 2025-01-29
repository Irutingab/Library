from utilities.loader import LibraryManagement
from utilities.user_edit import AddNewRowsandColumns
from utilities.Library_updater import UpdateSheet

def main():
    # Save data
    manager = LibraryManagement()
    manager.save_data_to_sheet()
    
    # Add new columns
    appender = AddNewRowsandColumns()
    appender.NewrowsAndcolumns()
    
    # Update rows
    updater = UpdateSheet()
    updater.SheetName()


if __name__ == "__main__":
    main()