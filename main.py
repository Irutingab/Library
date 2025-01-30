from utilities.loader import LibraryManagement

from utilities.user_edit import AddNewRowsandColumns

from utilities.Library_updater import UpdateSheet

from utilities.delete_row import DeleteRowFromFile

from utilities.colrecover import RecoverColumns

from utilities.Availablebooks import CountBooks


def main():
    # Save data
    manager = LibraryManagement()
    #manager.save_data_to_sheet()
    
    # Add new columns
    appender = AddNewRowsandColumns()
    #appender.NewrowsAndcolumns()
    
    # Update rows
    updater = UpdateSheet()
    updater.DatainCells()
    
    # Delete desired data or empty cells
    manager = DeleteRowFromFile()
    #manager.RemovecellValue()
    #manager.RemoveRow()
    #manager.RemoveCol()
    
    #to recover a deleted row
    manager = RecoverColumns()
    #manager.NewrowsAndcolumns()
    
    manager = CountBooks()  
    manager.AvailableBooks()


if __name__ == "__main__":
    main()