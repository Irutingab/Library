def display_menu():
    print("\nLibrary Management Menu:")
    print("1. Add New Book")
    print("2. Edit Existing Book")
    print("3. View All Books")  
    print("4. Save and Exit")
    print("5. Exit without Saving")
    choice = input("Enter your choice (1-5): ")
    return choice

def get_book_data():
    title = input("Enter book title: ")
    author = input("Enter author name: ")
    Publication_year = input("Enter pub_year: ")  
    return [title, author, Publication_year] #return a list: ordered, changeable, and allow duplicates

while True:
    choice = display_menu()
    if choice == '1':
        new_book = get_book_data()
        # ... (Code to add new_book to Excel)
    elif choice == '2':
        # ... (Code to handle editing)
    elif choice == '3':
        # ... (Code to view data)
    elif choice == '4':
        # ... (Save to Excel and break the loop)
        break
    elif choice == '5':
        break
    else:
        print("Invalid choice. Please try again.")