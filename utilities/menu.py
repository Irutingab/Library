def menu():
    print("1. Say Hello")
    print("2. Update")
    print("3. ")

while True:
    menu()
    choice = input("Enter your choice: ")
    if choice == "1":
        print("Hello!")
    elif choice == "2":
        print("Goodbye!")
    elif choice == "3":
        print("Exiting...")
        break
    else:
        print("Invalid choice. Try again.")
