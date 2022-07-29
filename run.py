
from find_info import find_info
from excel_read_write import excel_read_write
from ordered_set import OrderedSet
import re


def print_details(phones,emails):
    for phone in phones:
        print(phone)
    for email in emails:
        print(email)

#The main function that finds details
def func(browser,names,addresses,city_states):
    unit_addresses=filter(lambda x:x.find("unit")!=-1,addresses)
    for add in addresses:
        browser.land_first_page()
        browser.search(add)
        if(browser.check_address(names)):
            return(True)
        for name in names:
            for city_state in city_states:
                browser.land_first_page()
                browser.search(name+f" {city_state}")
                if(not browser.find_proper_name(name,addresses)):
                    if(unit_addresses):
                        browser.land_first_page()
                        browser.search(name+f" {city_state}")
                        if(browser.find_proper_name(name,unit_addresses,True)):
                            return(True)
                else:
                    return(True)
        return(False)

if __name__ == '__main__':

    browser=find_info.FindInfo()
    filepath=input("Enter the file path: ")
    savepath=input("Enter save file path: ")
    rw=excel_read_write.ExcelReadWrite(filepath,0)
    i=input("1.Select from existing profiles\n2.Create new profile\n")
    if(i=="1"):
        while(True):
            try:
                rw.set_profile()
                break
            except:
                print("Invalid profile name")
                
    elif(i=="2"):
        rw.create_profile()
    else:
        print("Invalid Option")
        exit()
    
    row_start=int(input("Enter starting row number: "))
    row_end=int(input("Enter ending row number: "))
    input("Press anything to start execution")
    for i in range(row_start-1,row_end):
        try:
            l=[[],[]]
            name_list=rw.fetch_names(i)
            name_list=list(filter(lambda a:a!="",name_list))
            print(name_list[0])
            unit_sep_addresses=rw.fetch_addresses(i)
            city_states=rw.fetch_city_states(i)
            units=rw.fetch_units(i)
            addresses=OrderedSet()
            for j in range(len(units)):
                if(city_states):
                    add=f"{unit_sep_addresses[j].lower()} {city_states[j].lower()}"
                else:
                    add=unit_sep_addresses[j].lower()
                if(units[j]!=-1):
                    add+=f" unit {units[j]}"
                addresses.add(add)
            city_states=OrderedSet(city_states)
            if(name_list):
                func(browser,name_list,addresses,city_states)
                l=browser.fetch_details()
                browser.clear()
            if(l[0] or l[1]):
                print("Found")
            else:
                print("Not Found")
            rw.write_phones(l[0],i)
            rw.write_emails(l[1],i)
        except:
            raise

    rw.save(savepath)
    print("Details Saved")
    input()