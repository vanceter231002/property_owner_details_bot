from tokenize import String
from types import NoneType
from openpyxl import load_workbook
from openpyxl import Workbook

class ExcelReadWrite:
        def __init__(self,filepath,sheet_no):
            self.wb=load_workbook(filepath)
            self.sheet=self.wb.worksheets[sheet_no]
            self.set_row_list()
            self.column_info={"fname_columns":[2],
                            "lname_columns":[3],
                            "address_columns":[4],
                            "city_columns":[5],
                            "state_columns":[6],
                            "phone_columns":"Default",
                            "email_columns":"Default",
                            "full_address":"n"}

        def set_init(self):
            print("Enter column letter for the following")
            n=input("Enter number of Names: ")
            n=int(n)
            l1=[]
            l2=[]
            for i in range(n):
                l1.append(self.get_column(input(f'First Name {i+1}: ')))
                l2.append(self.get_column(input(f"Last Name {i+1}: ")))
            self.column_info['fname_columns']=l1
            self.column_info['lname_columns']=l2
            self.column_info["full_address"]=input("Full address? (Y or N): ").lower()
            n=input("Enter number of Addresses: ")
            n=int(n)
            l1=[]
            l2=[]
            l3=[]
            for i in range(n):
                l1.append(self.get_column(input(f'Address {i+1}: ')))
                if(self.column_info["full_address"]=="n"):
                    l2.append(self.get_column(input(f'City {i+1}: ')))
                    l3.append(self.get_column(input(f'State {i+1}: ')))
            self.column_info['address_columns']=l1
            self.column_info['city_columns']=l2
            self.column_info['state_columns']=l3
            n=input("Enter number of phone numbers: ")
            n=int(n)
            l=[]
            for i in range(n):
                l.append(self.get_column(input(f'Phone {i+1}: ')))
            n=input("Enter number of Emails: ")
            n=int(n)
            self.column_info["phone_column"]=l
            l=[]
            for i in range(n):
                l.append(self.get_column(input(f'Email {i+1}: ')))
            self.column_info["email_column"]=l
            
        def get_column(self,string:str):
            index=0
            string=string[::-1]
            for i in range(len(string)):
                index+=(ord(string[i])-64)*26**i
            return(index-1)
        def no_of_rows(self):
            return(len(self.rows))
        def set_row_list(self):
            i=0
            for cell in self.sheet['E']:
                if(cell.value==None):
                    break
                i+=1
            self.rows=list(self.sheet.rows)[:i]

        def write_phones(self,phones,row_index):
            row=self.rows[row_index]
            j=0
            i=28
            for phone in phones:
                if(j==10):
                    break
                phone=phone.replace('(',"")
                phone=phone.replace(')',"")
                phone=phone.replace('-',"")
                phone=phone.replace(' ',"")
                row[i].value=phone
                i+=4
                j+=1
        def write_emails(self,emails,row_index):
            row=self.rows[row_index]
            j=0
            i=64
            for email in emails:
                if(j==5):
                    break
                row[i].value=email
                i+=3
                j+=1

        def fetch_names(self,row_index):
            row=self.rows[row_index]
            fname_columns=self.column_info["fname_columns"]
            lname_columns=self.column_info["lname_columns"]
            names=[]
            for i in range(len(fname_columns)):
                first_name=row[fname_columns[i]]
                last_name=row[lname_columns[i]]
                if(not first_name.value):
                    names.append("")
                elif(not last_name.value):
                    names.append(first_name.value.strip())
                else:
                    names.append(f"{first_name.value.strip()} {last_name.value.strip()}")
            return(names)        
        def fetch_city_state(self,row_index):
            state_city=f"{self.rows[row_index][self.get_column('F')].value.strip()} {self.rows[row_index][self.get_column('G')].value.strip()}"
            return(state_city)
        
        def fetch_address(self,row_index):
            column_no=self.get_column('E')
            address=self.rows[row_index][column_no].value.strip()
            i1=address.lower().find(" apt ")
            i2=address.lower().find(" unit ")
            if(i1!=-1):
                address=address[:i1]
            elif(i2!=-1):
                address=address[:i2]
            return(address.strip())
        
        def fetch_unit(self,row_index):
            address=self.rows[row_index][self.get_column('E')].value.strip()
            i1=address.lower().find(" apt ")
            i2=address.lower().find(" unit ")
            unit=-1
            if(i1!=-1):
                unit=address[i1+4:].strip()
            elif(i2!=-1):
                unit=address[i2+5:].strip()
            return(unit)
        def fetch_alt_address(self,row_index):
            column=self.get_column('CE')
            address=self.rows[row_index][column].value.strip()
            i1=address.lower().find(" apt ")
            i2=address.lower().find(" unit ")
            if(i1!=-1):
                address=address[:i1]
            elif(i2!=-1):
                address=address[:i2]
            return(address.strip())
        
        def fetch_alt_city_state(self,row_index):
            state_city=f"{self.rows[row_index][self.get_column('CF')].value.strip()} {self.rows[row_index][self.get_column('CG')].value.strip()}"
            return(state_city)
        def fetch_alt_unit(self,row_index):
            address=self.rows[row_index][self.get_column('CE')].value.strip()
            i1=address.lower().find(" apt ")
            i2=address.lower().find(" unit ")
            unit=""
            if(i1!=-1):
                unit=address[i1+4:].strip()
            elif(i2!=-1):
                unit=address[i2+5:].strip()
            return(unit)

        def save(self,filepath):
            self.wb.save(filepath)
        




            
            

