# student management system

# admin , teacher
# admin - login,add teacher , delete teachers ,log out
# teacher -  view all students , search a student , assign a task to student ,validate task , log out
from openpyxl import Workbook
from openpyxl import load_workbook
# from abc import ABC , abstractmethod


class admin():
    admin_username="admin"
    admin_password="password"
    def __init__(self,ausername,apassword):
        self.ausername=ausername
        self.apassword=apassword
    # @abstractmethod
    def login(self,i):
        while True:
                if i<3:
                    if self.ausername==self.admin_username:
                        if self.apassword==self.admin_password:
                            print("Admin login sucessfull")
                            return True 
                        else:
                            i+=1
                            print("password is invalid try again, remaining attempts", i,"/3")
                            self.ausername = input("Enter your username: ")
                            self.apassword=input(" your password : ")
                        
                        
                    else:
                        i+=1
                        print("Username is invalid try again" ,i,"/3")
                        self.ausername=input("Enter your username : ")
                        self.apassword=input(" your password : ")
                else:
                
                    print("you have exceeded all login attempts try again later ")
                    print("Login failed")
                    return False
    # @abstractmethod
    def new_user(self):
        try:
            f = load_workbook("teacher_info.xlsx")
            choice=input("File exists already do you want to continue by overwriting the existing file ? (yes or no)")
            while choice.lower() != "yes" and choice.lower() != "no":
                print("Please enter Yes or no")
                choice=input("Do you want to create the file by overwriting the old file ?(yes or no)")
            if choice.lower()=="yes":
                f = Workbook()
                ws = f.active
                ws["A1"]="Name"
                ws["B1"] = "UserName"
                ws["C1"] = "Password"
                ws["D1"]="department"
                ws.cell(row=2, column=1).value= "New"
                ws.cell(row=2, column=2).value = "New1"
                ws.cell(row=2, column=3).value = "password"
                ws.cell(row=2, column=4).value = "test"
                f.save("teacher_info.xlsx")
                print("The new file has  been created by overwriting the old file")
                
            else:
                print("File creation has been cancelled")

        except FileNotFoundError:   
            f = Workbook()
            ws = f.active
            ws["A1"]="Name"
            ws["B1"] = "UserName"
            ws["C1"] = "Password"
            ws["D1"]="department"
            ws.cell(row=2, column=1).value= "New"
            ws.cell(row=2, column=2).value = "New1"
            ws.cell(row=2, column=3).value = "password"
            ws.cell(row=2, column=4).value = "test"
            f.save("teacher_info.xlsx")
            print("The new file has  been created ")
    # @abstractmethod
    def menu(self,skip_ask=False):
        if skip_ask==True:
            choice="yes"
        else:
            choice=input("would you like to go back to admin menu ? : ").strip().lower()
            while choice!="yes" and choice!="no":
                print("Enther yes or no")
                choice=str(input("would you like to go back to admin menu ? :")).lower().strip()
        if choice=="yes":
            while True:
                try:
                    choice = int(input("==========Admin Menu=========="
                        "\nDo you want to add new teachers and credentials (1) : "
                        "\nDo you want to edit an existing teacher credentials (2) : "
                        "\nDo you want to delete a teacher's credential (3) : "
                        "\nDo you want to view the credential of an existing teacher (4) : "
                        "\nLogout (5) : "
                    ))
                    if 1 <= choice <= 5:
                        return choice
                    print("Please enter a number between 1 and 5.")
                except ValueError:
                    print("That's not a number, try again.")
        else:
            return 5

    def space_check(self,prompt):
        a=input(prompt).strip()
        while a =="":
            print("This field can't be empty, please try again.")
            a=input(prompt).strip()
        return a
    def edit_teacher(self):
        f = load_workbook("teacher_info.xlsx")
        ws = f.active
        choice=self.menu(skip_ask=True)
        while True:
            if choice == 1:
                name=self.space_check("Enter the name of the Teacher : ")
                username=input("Enter the username :  ")
                while True:
                    if username == "":
                        print("Please enter a valid input")
                    elif any(str(ws.cell(row=i, column=2).value) == username for i in range(2, ws.max_row + 1)):
                        print("This username already exists. Please enter another one.")
                    else:
                        break
                    username=input("Enter the username :  ")
                password=self.space_check("Enter the password : ")
                department=self.space_check("Enter the department : ")
                new_row=ws.max_row+1
                ws.cell(row=new_row, column=1).value= name
                ws.cell(row=new_row, column=2).value = username
                ws.cell(row=new_row, column=3).value = password
                ws.cell(row=new_row, column=4).value = department
                f.save("teacher_info.xlsx")
                print("The new teacher record has been recorded ")
                choice=self.menu()

            elif choice ==2:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the teacher ").strip().lower()
                        editdept=input("Enter the department ").strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i,column=1).value or "").lower().strip() ==editname and str(ws.cell(row=i,column=4 ).value or "").strip().lower()== editdept:
                                    found=True
                                    name=self.space_check("Enter the new Teachers name : ")
                                    username=input("Enter the new username :  ")
                                    while True:
                                        if username == "":
                                            print("Please enter a valid input")
                                        elif any(str(ws.cell(row=j, column=2).value) == username for j in range(2, ws.max_row + 1)  if j!=i) :
                                            print("This username already exists. Please enter another one.")
                                        else:
                                            break
                                        username=input("Enter the username :  ")
                                    password=self.space_check("Enter the  new password : ")
                                    department=self.space_check("Enter the new department : ")
                                    ws.cell(row=i, column=1).value= name
                                    ws.cell(row=i, column=2).value = username
                                    ws.cell(row=i, column=3).value = password
                                    ws.cell(row=i, column=4).value = department
                                    f.save("teacher_info.xlsx")
                                    print("The values has been edited ")
                                    choice=self.menu()
                                    break
                        if found==False:
                            a = input("Invalid input please enter valid teacher name and department , Would you like to try again ? (Yes or no) : ").strip().lower()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").lower().strip()
                        else:
                            break
                    
                    
                    else:
                         choice=self.menu()
                         break
            elif choice==3:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the teacher who you want to delete from the database  ").strip().lower()
                        editdept=input("Enter the department ").strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i,column=1).value or "").strip().lower() == editname and str(ws.cell(row=i,column=4).value or "").strip().lower() == editdept:
                                    found=True
                                    confirmation=input("Are you sure you want to delete the credentials ? (Yes or no)").strip()
                                    while confirmation.lower() != "yes" and confirmation.lower() != "no":  
                                        print("Please enter Yes or No")  
                                        confirmation = input("Are you sure you want to delete the credentials? (Yes or no): ").strip()
                                    if confirmation.lower()=="yes":
                                        ws.delete_rows(i, 1)
                                        f.save("teacher_info.xlsx")
                                        print("The selected teachers credentials has been deleted  ")
                                        choice=self.menu()
                                        break
                                    elif confirmation.lower() == "no":
                                        print("Deletion cancelled")
                                        choice=self.menu()
                                        break
                        if found==False:
                            a = input("Invalid input please enter valid teacher name and department , Would you like to try again ? (Yes or no) : ").strip()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").strip()
                        else:
                             break
                    else:
                         choice=self.menu()
                         break
            elif choice==4:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the teacher who's details you want to view from the database  ").lower().strip()
                        editdept=input("Enter the department ").lower().strip()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i,column=1).value or "").strip().lower() == editname and str(ws.cell(row=i,column=4).value or "").strip().lower() == editdept:
                                    found=True
                                    name = ws.cell(row=i, column=1).value  
                                    username = ws.cell(row=i, column=2).value  
                                    password = ws.cell(row=i, column=3).value  
                                    department = ws.cell(row=i, column=4).value  
                                    print(f"| Name: {name} | Username: {username} | Password: {password} | Department: {department} |")
                                    choice=self.menu()
                                    break
                        if found==False:
                            a = input("Invalid input please enter valid teacher name and department , Would you like to try again ? (Yes or no) : ").strip()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").strip().lower()
                        else:
                             break
                    else:
                        choice=self.menu()
                        break
            elif choice==5:
                break
            



                                    

                            


    
        
    

    


        