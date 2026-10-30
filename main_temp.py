# dummy temp main - imports admin.py / teacher.py as they are
import os
from admin import admin
from teacher import teacher

def ensure_files():
    if not os.path.exists("teacher_info.xlsx"):
        admin("", "").new_user()
    if not os.path.exists("student_info.xlsx"):
        teacher("", "").new_user()

def main():
    ensure_files()
    while True:
        role = input("\n1) Admin  2) Teacher  3) Exit : ").strip()
        if role == "1":
            a = admin(input("Username: "), input("Password: "))
            if a.login(0) == "Admin login sucessfull":
                print("Admin login sucessfull")
                a.edit_teacher()
        elif role == "2":
            t = teacher(input("Username: "), input("Password: "))
            if t.login(0) == "Teacher login sucessfull":
                print("Teacher login sucessfull")
                t.student_edit()
        elif role == "3":
            break
        else:
            print("Enter 1, 2 or 3.")

if __name__ == "__main__":
    main()
