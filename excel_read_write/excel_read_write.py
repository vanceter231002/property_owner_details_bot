from openpyxl import load_workbook
import json
import re

class ExcelReadWrite:
        def __init__(self,filepath,sheet_no):
            self.wb=load_workbook(filepath)
            self.sheet=self.wb.worksheets[sheet_no]
            self.rows=list(self.sheet.rows)
            f=open("excel_read_write/profiles.json")
            self.profiles=json.load(f)
            f.close()
            self.column_info=self.profiles["Default"]

        def set_profile(self):
            name=input("Enter name of the profile: ")
            self.column_info=self.profiles[name]
        
        def print_profile_list(self):
            for profile in self.profiles["names"]:
                print(profile)
            print("")
        
        def create_profile(self):
            name=input("Enter the name of profile: ")
            while(name in self.profiles["names"]):
                print("Profile name already in use")
                name=input("Enter the name of profile: ")
            d={}
            print("Enter column letter for the following: ")
            n=input("Enter number of Names: ")
            n=int(n)
            l1=[]
            l2=[]
            for i in range(n):
                l1.append(self.get_column(input(f'First Name {i+1}: ')))
                l2.append(self.get_column(input(f"Last Name {i+1}: ")))
            d['fname_columns']=l1
            d['lname_columns']=l2
            n=input("Enter number of Addresses: ")
            n=int(n)
            l1=[]
            l2=[]
            l3=[]
            for i in range(n):
                l1.append([self.get_column(input(f'Address {i+1}: ')),-1])
                x=input("Unit column?(Y or N)").lower()
                if(x=="y"):
                    l1[i][1]=self.get_column(input("Unit Column: "))
                l2.append(self.get_column(input(f'City {i+1}: ')))
                l3.append(self.get_column(input(f'State {i+1}: ')))
            d['address_columns']=l1
            d['city_columns']=l2
            d['state_columns']=l3
            n=input("Enter number of phone numbers: ")
            n=int(n)
            l=[]
            for i in range(n):
                l.append(self.get_column(input(f'Phone {i+1}: ')))
            n=input("Enter number of Emails: ")
            n=int(n)
            d["phone_columns"]=l
            l=[]
            for i in range(n):
                l.append(self.get_column(input(f'Email {i+1}: ')))
            d["email_columns"]=l
            self.column_info=d
            self.profiles[name]=d
            self.profiles["names"].append(name)
            f=open("excel_read_write/profiles.json","w")
            json.dump(self.profiles,f,indent=2)

        def get_column(self,string:str):
            string=string.upper()
            index=0
            string=string[::-1]
            for i in range(len(string)):
                index+=(ord(string[i])-64)*26**i
            return(index-1)

        def no_of_rows(self):
            return(len(self.rows))


        def write_phones(self,phones,row_index):
            row=self.rows[row_index]
            i=0
            l=self.column_info["phone_columns"]
            for phone in phones:
                if(i==len(l)):
                    break
                phone=phone.replace('(',"")
                phone=phone.replace(')',"")
                phone=phone.replace('-',"")
                phone=phone.replace(' ',"")
                row[l[i]].value=phone
                i+=1
                
                
        def write_emails(self,emails,row_index):
            row=self.rows[row_index]
            l=self.column_info["email_columns"]
            i=0
            for email in emails:
                if(i==len(l)):
                    break
                row[l[i]].value=email
                i+=1
                

        def fetch_names(self,row_index):
            row=self.rows[row_index]
            fname_columns=self.column_info["fname_columns"]
            lname_columns=self.column_info["lname_columns"]
            names=[]
            for i in range(len(fname_columns)):
                first_name=row[fname_columns[i]]
                last_name=row[lname_columns[i]]
                if(not first_name.value):
                    pass
                elif(not last_name.value):
                    names.append(first_name.value.strip())
                else:
                    names.append(f"{first_name.value.strip()} {last_name.value.strip()}")
            return(names)        

        def fetch_city_states(self,row_index):
            city_states=[]
            city_columns=self.column_info["city_columns"]
            state_columns=self.column_info["state_columns"]
            for i in range(len(city_columns)):
                city_state=f"{self.rows[row_index][city_columns[i]].value.strip()} {self.rows[row_index][state_columns[i]].value.strip()}"
                city_states.append(city_state)
            return(city_states)
        
        def fetch_addresses(self,row_index):
            addresses=[]
            address_columns=self.column_info["address_columns"]
            for column in address_columns:
                address=self.rows[row_index][column[0]].value.strip()
                if(address):
                    i1=address.lower().find(" apt ")
                    i2=address.lower().find(" unit ")
                    if(i1!=-1):
                        address=address[:i1]
                    elif(i2!=-1):
                        address=address[:i2]
                    addresses.append(address.strip())
            return(addresses)
        
        def fetch_units(self,row_index):
            units=[]
            for column in self.column_info["address_columns"]:
                if(column[1]!=-1):
                    unit=self.rows[row_index][column[1]].value.strip()
                    if(unit):
                        unit=re.search("[0-9a-zA-Z]+",unit).group()
                    else:
                        unit=-1
                    units.append(unit)
                else:
                    address=self.rows[row_index][column[0]].value.strip()
                    if(address):
                        i1=address.lower().find(" apt ")
                        i2=address.lower().find(" unit ")
                        unit=-1
                        if(i1!=-1):
                            unit=address[i1+4:].strip()
                        elif(i2!=-1):
                            unit=address[i2+5:].strip()
                        units.append(unit)
            return(units)

        def save(self,filepath):
            self.wb.save(filepath)
        




            
            

