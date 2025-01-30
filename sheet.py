"""
MODULE TO MANAGE SHEET OPERATIONS

"""
import mailer

def isValid(rollNumber, email):
    return rollNumber != "" and email != ""

def setupSheet(sheet, db):
        '''
        FUNCTION TO SETUP COLUMN IN SHEET AND SUBJECTS 
        '''
        if not sheet:
            print("Error: Worksheet is not initialized. Exiting setup.")
            return

        sheet['A1'] = "Roll Number"
        sheet['B1'] = "Student Email"

        try:
        
            num_subjects = int(input("Enter the number of subjects: "))
        
            if num_subjects <= 0:
                print("The number of subjects must be greater than zero.")
                return

            cols = ["C", "D", "E", "F", "G", "H"] 

            if num_subjects > len(cols):
                print("Maximum 6 subjects are supported in this implementation.")
                num_subjects = len(cols)

            for i in range(num_subjects):
        
                subject_name = input(f"Enter the name for Subject {i + 1}: ")
                sheet[f"{cols[i]}1"] = subject_name
            
            db.save("attendance.xlsx")

        except ValueError:
            print("Invalid input. Please enter a number.") 



def addStudent(sheet, db, validateSubject):
    """
    FUNCTION TO ADD NEW STUDENTS INTO SHEET
    """
    stduents = int(input("Number of students you want to add: "))
        
    row = sheet.max_row + 1

    for i in range(stduents):
        print("Enter student Roll Number and email id in format XXXXX testinform@test.com: ")
        roll_number, email = input("").split(" ")
            
        while not isValid(roll_number, email):
            print("Enter valid Roll Number and email in following format  XXXXX testinform@test.com: ")
            roll_number, email = input("").split(" ")
            
        if roll_number != '' and email != '':
            sheet.cell(row=row, column=1).value = roll_number
            sheet.cell(row=row, column=2).value = email
            for index, i in enumerate(validateSubject, start=3):
                sheet.cell(row=row, column=index).value = 0
            row += 1

    db.save("attendance.xlsx")



def saveAttendace(sheet, db, row, col):
        """
        FUNCTION TO SAVE ATTENDANEC
        """
        curr = sheet.cell(row=row, column=col).value
        sheet.cell(row=row, column=col).value = curr + 1
        db.save("attendance.xlsx")
     
        if  curr + 1 == 2 :
            mailer.oneMoreLeaveMessage(sheet.cell(row=row, column=2).value)

        if  curr + 1 == 3 :
            mailer.leaveQuotaExceedMail(sheet.cell(row=row, column=2).value)

       



def resetAttendance(sheet):
    """
    FUNCTION TO RESET SHEET ATTENDANCE
    """
    
    for i in range(1, len(sheet.max_row)+1):
            
        if sheet.max_column > 2:
           sheet.cell(row=i, column=3).value = 0     
            
        if sheet.max_column > 3:  
           sheet.cell(row=i, column=4).value = 0

        if sheet.max_column > 4:  
           sheet.cell(row=i, column=5).value = 0  

        if sheet.max_column > 5:  
           sheet.cell(row=i, column=6).value = 0    

        if sheet.max_column > 6:  
           sheet.cell(row=i, column=7).value = 0    

        if sheet.max_column > 7: 
           sheet.cell(row=i, column=8).value = 0    