import pandas as pd
excel_file = 'participant.xlsx'
data = pd.read_excel(excel_file) #reads the excel file
data=data.set_index("Sl.No") #Setting Column 1 as index
num=data["Participant ID"].tolist() #converts column to list
limit=len(num) #number of rows
while True :
        to_be_searched = input("Enter ID: ") #enter data to be searched
        if to_be_searched == "None" or to_be_searched == "none":
                print("Exiting Program")
                break
        else:
                index = None
                for i in range(limit):
                        if to_be_searched == num[i]:
                                index = i
                                data.loc[index+1, "Attendance"] = "Present"
                                print("Operation Successful")
                                print(data.loc[index+1],"\n")
                                break;

                if index == None:
                        print("Member not found \n")
                
data.to_excel("final.xlsx")
