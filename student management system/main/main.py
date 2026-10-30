# import os
from admin import admin
from teacher import teacher
from openpyxl import load_workbook
def main_menu(menu_skip=True):
    if menu_skip == False:
        choice = input("would you like to go back to the main menu ? : ").strip().lower()
        while choice != "yes" and choice != "no":
            print("Enter yes or no")
            choice = input("would you like to go back to the main menu ? : ").strip().lower()
        if choice == "no":
            return 5
    while True:
        try:
            choice = int(input("==========Welcome to the main menu=========="
                "\n 1) Admin menu "
                "\n 2) Teacher menu"
                "\n 3) New user database creation(Admin)"
                "\n 4) New user database creation(Teacher)"
                "\n 5) Exit : "))
            if 1 <= choice <= 5:
                return choice
            print("Please enter a number between 1 and 5.")
        except ValueError:
            print("That's not a number, try again.")
def menu_choice():
    choice=main_menu(menu_skip=True)
    while True:
        if choice==1:
            try :
                load_workbook("teacher_info.xlsx") 
            except FileNotFoundError:
                print("No database found please select option 3 first ")
            else:    
                username=input("Enter admin username : ")
                password=input("Enter the admin password : ")
                admin_main=admin(username,password)
                if admin_main.login(0):
                    print("Admin login sucessfull")
                    admin_main.edit_teacher()
                else:
                    print("Admin login failed")
        if choice==2:
            try :
                load_workbook("student_info.xlsx") 
            except FileNotFoundError:
                print("No database found please select option 3 first ")
            else:
                username=input("Enter teacher username : ")
                password=input("Enter the teacher password : ")
                teacher_main=teacher(username,password)
                if teacher_main.login(0):
                    print("Teacher login sucessfull")
                    teacher_main.student_edit()
                else:
                    print("Teacher login failed")
        if choice ==3:
            admin("", "").new_user()

        elif choice == 4:
            teacher("", "").new_user()

        elif choice == 5:
            print("Goodbye")
            break    
        choice = main_menu(menu_skip=False)
if __name__ == "__main__":
    menu_choice()



