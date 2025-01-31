from utilities.loader import LibraryManagement

from utilities.Library_updater import UpdateSheet


def main():
    # Save data
    manager = LibraryManagement()
    manager.save_data_to_sheet()
    
    #update data
    manager = UpdateSheet()
    manager.UpdateDatainCells()
    
if __name__ == "__main__":
    main()