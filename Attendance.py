import pandas as pd
excel_file = 'participant.xlsx'
data = pd.read_excel(excel_file) #reads the excel file
data=data.set_index("Sl.No") #Setting Column 1 as index
num=data["Participant ID"].tolist() #converts column to list
limit=len(num) #number of rows
p=0 #initialise number of present participant as 0
while True :
        to_be_searched = input("Enter ID: ") #enter data to be searched
        if to_be_searched == "None" or to_be_searched == "none":
                print("Exiting Program\n\n")
                print("Present Members: ",p,"\nAbsent Members: ",limit-p,"\n")
                break
        
        if to_be_searched == "Status" or to_be_searched == "status":
                for i in range(20):
                        print("-",end="-")
                print("\n\nPresent Members: ",p,"\nAbsent Members: ",limit-p,"\n")
                for i in range(20):
                        print("-",end="-")
                print("\n")
                
        else:
                index = None
                for i in range(limit):
                        if to_be_searched == num[i]:
                                index = i
                                data.loc[index+1, "Attendance"] = "Present"
                                p=p+1
                                print("Operation Successful")
                                print(data.loc[index+1],"\n")
                                break;

                if index == None:
                        print("Member not found \n")
                
data.to_excel("final.xlsx")
