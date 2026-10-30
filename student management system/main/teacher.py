# teacher -  view all students , search a student , assign a task to student ,validate task , log out
from openpyxl import Workbook
from openpyxl import load_workbook

class teacher:
    def __init__(self,tusername,tpassword):
        self.tusername=tusername
        self.tpassword=tpassword



    def login(self,j):
        f = load_workbook("teacher_info.xlsx")
        ws = f.active
        while True:
            if j<3:
                for i in range (2,ws.max_row+1):
                    if ws.cell(row=i,column=2).value ==self.tusername and ws.cell(row=i,column=3).value== self.tpassword:
                        print("Teacher login sucessfull") 
                        return True            
                j+=1
                print("Username or password is invalid try again" ,j,"/3")
                self.tusername=input("Enter your username : ")
                self.tpassword=input(" your password : ")
            else:
                
                print("you have exceeded all login attempts try again later ")
                print("Login failed")
                return False

            
    def new_user(self):
        try:
            f = load_workbook("student_info.xlsx")
            choice=input("File exists already do you want to continue by overwriting the existing file ? (yes or no)")
            while choice.lower() != "yes" and choice.lower() != "no":
                print("Please enter Yes or no")
                choice=input("Do you want to create the file by overwriting the old file ?(yes or no)")
            if choice.lower()=="yes":
                f = Workbook()
                ws = f.active
                ws["A1"]="Name"
                ws["B1"] = "Roll Number"
                ws["C1"] = "Department"
                ws["D1"]="Task"
                ws["E1"]= "Task status"
                f.save("student_info.xlsx")
                print("The new file has  been created by overwriting the old file"
                      "\n If you are new and dont have a credential use the following :"
                      "\n Username: new1"
                      "\n Password: password")
            else:
                print("File creation has been cancelled")
        except FileNotFoundError:   
                    f = Workbook()
                    ws = f.active
                    ws["A1"]="Name"
                    ws["B1"] = "Roll Number"
                    ws["C1"] = "Department"
                    ws["D1"]="Task"
                    ws["E1"]= "Task status"
                    f.save("student_info.xlsx")
                    print("The new file has  been created "
                        "\n If you are new and dont have a credential use the following :"
                        "\n Username: new1"
                        "\n Password: password")



    def menu(self,skip_ask=False):
        if skip_ask==True:
                choice="yes"
        else:
            choice=input("would you like to go back to teacher menu ? : ").strip().lower()
            while choice!="yes" and choice!="no":
                print("Enther yes or no")
                choice=str(input("would you like to go back to teacher menu ? :")).lower().strip()
        if choice=="yes":
            while True:
                try:
                    choice = int(input("==========Teacher Menu=========="
                        "\nDo you want to add new student details (1) : "
                        "\nDo you want to edit an existing studnet details (2) : "
                        "\nDo you want to delete a student details (3) : "
                        "\nDo you want to view the details of an existing student (4) : " \
                        "\nDo you want to view the details of all students (5) : "
                        "\nDo you want to assign a task to a student (6) : " 
                        "\nDo you want to validate the task given (7) : " 
                        "\nLogout (8) :"
                    ))
                    if 1 <= choice <= 8:
                        return choice
                    print("Please enter a number between 1 and 8.")
                except ValueError:
                    print("That's not a number, try again.")
        else:
             return 8
        

    def student_edit(self):
        f = load_workbook("student_info.xlsx")
        ws = f.active
        choice=self.menu(skip_ask=True)   
        while True:
            if choice == 1:
                name=input("Enter the name of the student : ").strip()
                while name =="":
                     print("Please enter a valid student name ")
                     name=input("Enter the name of the student : ").strip()
                rollnumber = input("Enter the rollnumber: ").strip()
                while True:
                    if rollnumber == "":
                        print("Please enter a valid input")
                    elif any(str(ws.cell(row=i, column=2).value) == rollnumber for i in range(2, ws.max_row + 1)):
                        print("This roll number already exists. Please enter another one.")
                    else:
                        break
                    rollnumber = input("Enter the rollnumber: ").strip()
                department=input("Enter the department : ").strip()
                while department=="":
                    print("Please enter a valid input")
                    department=input("Enter the department : ").strip()            
                ask_task = input("Do you have a task assigned for this student? ").strip().lower()
                while ask_task not in ["yes", "no"]:
                    print("Please enter yes or no")
                    ask_task = input("Do you have a task assigned for this student? ").strip().lower()
                if ask_task == "yes":
                    task = input("Enter the task: ").strip()
                    while task == "":
                        print("Please enter a valid task")
                        task = input("Enter the task: ").strip()
                    taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                    while taskval not in["complete","incomplete"]:
                        print("Please enter a valid input")
                        taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                if ask_task == "no":
                    task = "nil"
                    taskval="nil"
                    print("No task assigned. Task and task validity status has been set to nil.")
                new_row=ws.max_row+1
                ws.cell(row=new_row, column=1).value= name
                ws.cell(row=new_row, column=2).value = rollnumber
                ws.cell(row=new_row, column=3).value = department
                ws.cell(row=new_row, column=4).value = task
                ws.cell(row=new_row,column=5).value=taskval
                f.save("student_info.xlsx")
                print("New student details has been stored")
                choice=self.menu()
            elif choice==2:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the student ").strip().lower()
                        editroll=str(input("Enter the rollnumber of the student ")).strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i, column=1).value or "").strip().lower() == editname and str(ws.cell(row=i, column=2).value or "").strip().lower() == editroll:
                                    found=True
                                    name=input("Enter the name of the student : ").strip()
                                    while name =="":
                                            print("Please enter a valid student name ")
                                            name=input("Enter the name of the student : ").strip()
                                    rollnumber = input("Enter the rollnumber: ").strip()
                                    while True:
                                        if rollnumber == "":
                                            print("Please enter a valid input")
                                        elif any(str(ws.cell(row=j, column=2).value) == rollnumber for j in range(2, ws.max_row + 1) if j != i):
                                            print("This roll number already exists. Please enter another one.")
                                        else:
                                            break
                                        rollnumber = str(input("Enter the rollnumber: ")).strip()
                                    department=input("Enter the department : ").strip()
                                    while department=="":
                                        print("Please enter a valid input")
                                        department=input("Enter the department : ").strip()            
                                    ask_task = input("Do you have a task assigned for this student? ").strip().lower()
                                    while ask_task not in ["yes", "no"]:
                                        print("Please enter yes or no")
                                        ask_task = input("Do you have a task assigned for this student? ").strip().lower()
                                    if ask_task == "yes":
                                        task = input("Enter the task: ").strip()
                                        while task == "":
                                            print("Please enter a valid task")
                                            task = input("Enter the task: ").strip()
                                        taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                                        while taskval not in["complete","incomplete"]:
                                            print("Please enter a valid input")
                                            taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                                    if ask_task == "no":
                                        task = "nil"
                                        taskval="nil"
                                        print("No task assigned. Task and task validity status has been set to nil.")
                                    ws.cell(row=i, column=1).value= name
                                    ws.cell(row=i, column=2).value = rollnumber
                                    ws.cell(row=i, column=3).value = department
                                    ws.cell(row=i, column=4).value = task
                                    ws.cell(row=i,column=5).value=taskval
                                    f.save("student_info.xlsx")
                                    print("The values has been edited ")
                                    choice = self.menu()
                                    break
                        if found == False:
                            a = input("Invalid input please enter valid student name and rollnumber, Would you like to try again? (Yes or no): ").strip()
                            while a.lower() != "yes" and a.lower() != "no":
                                print("Please enter Yes or No")
                                a = input("Would you like to try again? (Yes or no): ").strip()
                        else:
                            break
                        if a.lower() == "no":
                            choice = self.menu()
                            break
                    else:
                        choice=self.menu()
                        break
            elif choice==3:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the student who you want to delete from the database  ").strip().lower()
                        editroll=input("Enter the rollnumber ").strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                            if str(ws.cell(row=i,column=1).value or "").strip().lower() == editname and str(ws.cell(row=i,column=2).value or "").strip().lower() == editroll:
                                found=True
                                confirmation=input("Are you sure you want to delete the credentials ? (Yes or no)").strip()
                                while confirmation.lower() != "yes" and confirmation.lower() != "no":  
                                    print("Please enter Yes or No")  
                                    confirmation = input("Are you sure you want to delete the details? (Yes or no): ").strip()
                                if confirmation.lower()=="yes":
                                    ws.delete_rows(i, 1)
                                    f.save("student_info.xlsx")
                                    print("The selected student details has been deleted  ")
                                    choice=self.menu()
                                    break
                                elif confirmation.lower() == "no":
                                    print("Deletion cancelled")
                                    choice=self.menu()
                                    break
                        if found==False:
                            a = input("Invalid input please enter valid student name and rollnumber , Would you like to try again ? (Yes or no) : ").strip()
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
                        editname=input("Enter the name of the student who's details you want to view from the database : ").lower().strip()
                        editroll=input("Enter the roll number :  ").lower().strip()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i,column=1).value or "").strip().lower() == editname and str(ws.cell(row=i,column=2).value or "").strip().lower() == editroll:
                                    found=True
                                    name=ws.cell(row=i, column=1).value
                                    rollnumber=ws.cell(row=i, column=2).value
                                    department=ws.cell(row=i, column=3).value
                                    task=ws.cell(row=i, column=4).value
                                    taskval=ws.cell(row=i,column=5).value 
                                    print(f"| Name: {name} | Roll Number: {rollnumber} | Department: {department} | Task: {task} | Task status: {taskval} |")
                                    choice=self.menu()
                                    break
                        if found==False:
                            a = input("Invalid input please enter valid student name and roll number , Would you like to try again ? (Yes or no) : ").strip()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").strip().lower()
                        else:
                                break
                    else:
                        choice=self.menu()
                        break
            elif choice==5:
                if ws.max_row < 2:
                    print("No students found.")
                else:
                    for i in range(2, ws.max_row+1):
                        name=ws.cell(row=i, column=1).value
                        rollnumber=ws.cell(row=i, column=2).value
                        department=ws.cell(row=i, column=3).value
                        task=ws.cell(row=i, column=4).value
                        taskval=ws.cell(row=i, column=5).value
                        print(f"| Name: {name} | Roll Number: {rollnumber} | Department: {department} | Task: {task} | Task status: {taskval} |")
                choice=self.menu()
            elif choice==6:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the student ").strip().lower()
                        editroll=str(input("Enter the rollnumber of the student ")).strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i, column=1).value or "").strip().lower() == editname and str(ws.cell(row=i, column=2).value or "").strip().lower() == editroll:
                                    found=True
                                    ask_task = input("Would you like to assign a task for this student? ").strip().lower()
                                    while ask_task not in ["yes", "no"]:
                                        print("Please enter yes or no")
                                        ask_task = input("would you like to assign a task for this student? ").strip().lower()
                                    if ask_task == "yes":
                                        task = input("Enter the task: ").strip()
                                        while task == "":
                                            print("Please enter a valid task")
                                            task = input("Enter the task: ").strip()
                                        taskval="incomplete"  
                                        print("Task has been sucessfully assigned")                   
                                    elif ask_task == "no":
                                        task = "nil"
                                        taskval="nil"
                                        print("No task assigned. Task and task validity status has been set to nil.")
                                    ws.cell(row=i, column=4).value = task
                                    ws.cell(row=i,column=5).value=taskval
                                    f.save("student_info.xlsx")
                                    choice = self.menu()
                                    break
                        if found==False:
                            a = input("Invalid input please enter valid student name and roll number , Would you like to try again ? (Yes or no) : ").strip()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").strip().lower()
                        else:
                                break
                    else:
                        choice=self.menu()
                        break
            elif choice==7:
                a="Yes"
                while True:
                    if a.lower()=="yes":
                        editname=input("Enter the name of the student ").strip().lower()
                        editroll=str(input("Enter the rollnumber of the student ")).strip().lower()
                        found=False
                        for i in range (2,ws.max_row+1):
                                if str(ws.cell(row=i, column=1).value or "").strip().lower() == editname and str(ws.cell(row=i, column=2).value or "").strip().lower() == editroll:
                                    found=True
                                    name=ws.cell(row=i, column=1).value
                                    rollnumber=ws.cell(row=i, column=2).value
                                    department=ws.cell(row=i, column=3).value
                                    task=ws.cell(row=i, column=4).value
                                    taskval=ws.cell(row=i,column=5).value 
                                    print(f"| Name: {name} | Roll Number: {rollnumber} | Department: {department} | Task: {task} | Task status: {taskval} |")
                                    choice=str(input("would you like to change the task staus for this student and  task ? ")).lower().strip()
                                    while choice!="yes" and choice!="no":
                                        print("Enther yes or no")
                                        choice=str(input("would you like to change the task staus for this student and  task ? ")).lower().strip()
                                    if choice=="yes":      
                                        taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                                        while taskval not in["complete","incomplete"]:
                                            print("Please enter a valid input")
                                            taskval=input("Enter validity status of the task (complete/incomplete)").strip().lower()
                                        ws.cell(row=i,column=5).value=taskval
                                        f.save("student_info.xlsx")
                                        print("====================The task has been validated====================")
                                        print(f"| Name: {name} | Roll Number: {rollnumber} | Department: {department} | Task: {task} | Task status: {taskval} |")
                                        choice=self.menu()
                                        break
                                    elif choice=="no":
                                        print("Task validation cancelled.")
                                        choice=self.menu()
                                        break
                        if found==False:
                            a = input("Invalid input please enter valid student name and roll number , Would you like to try again ? (Yes or no) : ").strip()
                            while a.lower() != "yes" and a.lower() != "no":  
                                print("Please enter Yes or No")  
                                a = input("Would you like to try again? (Yes or no): ").strip().lower()
                        else:
                                break
                    else:
                        choice=self.menu()
                        break
            elif choice == 8:
                print("Teacher logged out successfully.")
                break
                                        
