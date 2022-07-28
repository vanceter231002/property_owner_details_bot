
from find_info import find_info
from find_info import excel_read_write
from ordered_set import OrderedSet


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
    rw=excel_read_write.ExcelReadWrite("../Files/copy.xlsx",0)
    browser=find_info.FindInfo()

    for i in range(186,190):
        try:
            l=[[],[]]
            name_list=rw.fetch_names(i)
            name_list=list(filter(lambda a:a!="",name_list))
            print(name_list[0])
            addresses=OrderedSet()
            add=f"{rw.fetch_address(i).lower()} {rw.fetch_city_state(i).lower()}"
            if(rw.fetch_unit(i)!=""):
                add+=f" unit {rw.fetch_unit(i)}"
            addresses.add(add)
            alt_add=f"{rw.fetch_alt_address(i).lower()} {rw.fetch_alt_city_state(i).lower()}"
            if(rw.fetch_alt_unit(i)!=""):
                alt_add+=f" unit {rw.fetch_alt_unit(i)}"
            addresses.add(alt_add)
            city_states=OrderedSet()
            city_state=rw.fetch_city_state(i)
            alt_city_state=rw.fetch_alt_city_state(i)
            city_states.add(city_state.lower())
            city_states.add(alt_city_state.lower())
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

    rw.save("../Files/copy.xlsx")
    print("Details Saved")
    input()