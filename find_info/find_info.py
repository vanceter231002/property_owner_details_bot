import undetected_chromedriver as uc
from selenium.webdriver.chromium.options import ChromiumOptions
import os
from selenium.webdriver.common.by import By
import time
import random



uc.TARGET_VERSION=104
class FindInfo(uc.Chrome):
    def __init__(self,userdata_dir=r"--user-data-dir=C:\Users\vansh\AppData\Local\Google\Chrome Beta\User Data",teardown=False,profile="Profile 1"):
            self.teardown=teardown
            self.phones=set()
            self.emails=set()
            os.environ['PATH']+=f";{self.driver_path}"
            #if options=="nolog":
            options=ChromiumOptions()
            options.add_argument(f"--user-data-dir={userdata_dir}")
            options.add_argument(fr'--profile-directory={profile}')
            #options.add_argument("--no-first-run")
            #options.add_argument("--log-level=3")
            #options.add_experimental_option('excludeSwitches',['enable-logging'])
            options.add_argument("--window-size=1366,768")
            #options.add_argument("user-agent= ")
            #options.binary_location="C:\Program Files\Google\Chrome Beta\Application\chrome.exe"
            super().__init__(options=options)
            self.implicitly_wait(15)
            self.times=[1,1,1.5,2]

    def __exit__(self,exc_type,exc_val,exc_tb):
        if self.teardown:
            self.quit()
    
    def get(self,url):
        time.sleep(random.choice(self.times))
        super().get(url)

    def clear(self):
        self.phones.clear()
        self.emails.clear()

    #land first page
    def land_first_page(self):
        self.get('https://www.spokeo.com/')
    
    #search a name or address
    def search(self,inp):
        search_form=self.find_element(By.CSS_SELECTOR,'form[data-testid="hero-search-box"]')
        select_button=search_form.find_element(By.CSS_SELECTOR,'input[aria-label="Search"]')
        select_button.send_keys(inp)
        final_button=search_form.find_element(By.CSS_SELECTOR,'button[id="search"]')
        final_button.click()
    
    #find if a address contains the required name in the list of its residents
    def check_address(self,names):
        if(self.check_function(By.CSS_SELECTOR,'ul[aria-label="Breadcrumbs"]',time=5)):
            self.implicitly_wait(0)
            people=self.find_elements(By.XPATH,f'(//div[@id="property-owners"]|//div[@id="current-residents"]|//div[@id="previous-residents"])//div[@class="content"]/a')
            check_names=[]
            links=[]
            for person in people:
                temp1=person.get_attribute("textContent")
                temp2=person.get_attribute("href")
                if(temp2 not in links):
                    check_names.append(temp1)
                    links.append(temp2)
            no_of_people=len(check_names)
        else:
            print("Some error while checking address")
            return(False)
        self.implicitly_wait(15)
        for name in names:
            for i in range(no_of_people):
                person_text=check_names[i]
                first_name=person_text.lower().split()[0]
                last_name=person_text.lower().split()[-1]
                check_first_name=name.lower().split()[0]
                if(self.verify_names(person_text.lower(),name.lower())):
                    temp=len(links[i])-links[i][::-1].find('/')-1
                    link=f"{links[i][:temp+1]}t{links[i][temp+2:]}"
                    self.get(link)
                    if(self.get_details()):
                        return(True)
                elif(first_name==check_first_name or last_name==check_first_name):
                    self.get(links[i])
                    self.find_element(By.CSS_SELECTOR,"#summary-avatar-panel")
                    if(self.check_function(By.CSS_SELECTOR,"div#summary-description")):
                        aka=self.find_element(By.CSS_SELECTOR,"div#summary-description")
                        names=aka.text[5:].split(", ")
                    else:
                        names=[]
                    for n in names:
                        if(self.verify_names(n.lower(),name.lower())):
                            temp=len(links[i])-links[i][::-1].find('/')-1
                            link=f"{links[i][:temp+1]}t{links[i][temp+2:]}"
                            self.get(link)
                            if(self.get_details()):
                                return(True)
        return(False)
        #print("Name not found in address lookup") 
        #return(False)    

    #check for the person on a given address unit wise
    def check_unit_wise(self,name):
        unit_elements=self.find_elements(By.XPATH,'//div[contains(@class,"four-column-list-item")]//a[text()="VIEW DETAILS"]')
        links=[element.get_attribute("href") for element in unit_elements]
        for i in range(len(links)):
            self.get(links[i])
            self.check_address(name,True)
            self.back()
            if(len(self.phones)>=10 and len(self.emails)>=5):
                break
        
            

        

    #click on view details button of a name card
    # def view_details(self):
    #     self.find_element(By.CSS_SELECTOR,"div#locations-card")
    #     if(self.check_function(By.CSS_SELECTOR,'div#locations-card h2 a')):
    #         view_details=self.find_element(By.CSS_SELECTOR,"div#locations-card h2 a")
    #         view_details.click()
    #         return(True)
    #     return(False)

    #returns the phones and emails   
    def fetch_details(self):
        l=[self.phones.copy(),self.emails.copy()]
        return(l)

    #verify if 2 names are similar
    def verify_names(self,name1,name2):
        #print(name1,name2)
        l1=name1.split()
        l2=name2.split()
        if(l2[0] in l1 and l2[-1] in l1):
            return(True)
        return(False)

    #get details of a name card
    def get_details(self):
        if(len(self.phones)>=10 and len(self.emails)>=5):
            return(True)
        self.find_element(By.CSS_SELECTOR,"div#contacts-card")
        details=[]
        if(self.check_function(By.CSS_SELECTOR,"div#contacts-card div.trigger-wrapper h4")):
            details=[detail.text for detail in self.find_elements(By.CSS_SELECTOR,"div#contacts-card div.trigger-wrapper h4")]
        if(details):
            for detail in details:   
                if(detail.find('@')==-1):
                    self.phones.add(detail)
                else:
                    self.emails.add(detail)
            return(True)
        else:
            print("No contacts Found")
            return(False)
    def check_function(self,by,query,element=None,time=0):
        self.implicitly_wait(time)
        try:
            if element:
                element.find_element(by,query)
            else:
                self.find_element(by,query)
            self.implicitly_wait(15)
            return(True)
        except:
            self.implicitly_wait(15)
            return(False)

    #check if a search result name matches the orignal name
    def find_candidate_name(self,names,name):        
        name=name.lower()
        for check_name in names:
            if(self.verify_names(check_name.lower(),name)):
                return(True)
        return(False)

    #find from a list of name search results, which name contains the address
    def find_proper_name(self,name,addresses,flag=False):
        if(self.check_function(By.CSS_SELECTOR,'div[aria-label="Search Results"]',time=5)):
            if(self.check_function(By.XPATH,'(//div[@type="danger"]//strong)[2][contains(text(),"State")]')):
                return(-1)
            name_cards=[]
            if(self.check_function(By.CSS_SELECTOR,'a[aria_label="Page 1"]')):
                next=self.find_element(By.CSS_SELECTOR,'a[aria_label="Page 1"]').get_attribute("href")
                n=len(self.find_elements(By.CSS_SELECTOR,'div.pagination a'))-2
                name_cards=[]
                for i in range(1,n+1):
                    name_cards+=self.find_elements(By.CSS_SELECTOR,f'div[aria-label="Search Results"] .single-column-list-item')
                    temp=next.find("?")
                    if(i==2):
                        next=f"{next[:temp]}{i}{next[temp:]}"
                        self.get(next)
                    elif(i>2):
                        next=f"{next[:temp-1]}{i}{next[temp:]}"
                        self.get(next)
            else:
                name_cards=self.find_elements(By.CSS_SELECTOR,f'div[aria-label="Search Results"] .single-column-list-item')
        
            no=len(name_cards)
            name_list=[]
            links=[]
            for name_card in name_cards:
                l=[]
                link=name_card.find_element(By.CSS_SELECTOR,".button").get_attribute('href')
                temp=len(link)-link[::-1].find('/')-1
                link=f"{link[:temp+1]}t{link[temp+2:]}"
                link="https://www.spokeo.com"+link
                links.append(link)
                main_check_name=name_card.find_element(By.CSS_SELECTOR,'.title').get_attribute("textContent")
                temp=main_check_name.find(',')
                if(temp!=-1):
                    main_check_name=main_check_name[:temp]
                l.append(main_check_name)
                if(self.check_function(By.XPATH,'.//strong[text()="Also known as"]/following-sibling::span',name_card)):
                    aka_names=name_card.find_elements(By.XPATH,'.//strong[text()="Also known as"]/following-sibling::span')
                    for aka_name in aka_names:
                        l.append(aka_name.get_attribute("textContent"))
                name_list.append(l)
            for i in range(no):
                if(self.find_candidate_name(name_list[i],name)):
                    self.get(links[i])
                    for add in addresses:
                        if(self.check_name(add,flag)):
                            if(self.get_details()):
                                return(True)                
        return(False)
    #Verify if a following Name card has a address in address list
    def check_name(self,address,flag=False):
            d={"street":"st","road":"rd","avenue":"ave","east":"e","west":"w","south":"s","north":"n","drive":"dr"}
            i=address.lower().find("unit")
            if(i==-1):
                flag=True
                unit_sep_ad=address.lower()
            else:
                unit=address[i+5:].strip()
                address=address.lower()
                unit_sep_ad=address[:i-1].strip()
            
            for word in d.keys():
                unit_sep_ad=unit_sep_ad.replace(word,d[word])
            ver_addresses_main=[add.text for add in self.find_elements(By.CSS_SELECTOR,"div#locations-vertical-list div.profile-section-row div.tertiary-list-title")]
            ver_addresses_cs=[add.text for add in self.find_elements(By.CSS_SELECTOR,"div#locations-vertical-list div.profile-section-row div.subtitle-container")]
            for i in range(len(ver_addresses_main)):
                ver_address=ver_addresses_main[i]
                city_state=ver_addresses_cs[i]
                comma=city_state.find(',')
                city_state=city_state[:comma]+" "+city_state[comma+1:].strip().split()[0]
                comma=ver_address.find(',')
                ver_unit=ver_address[comma+1:].strip().split()[-1]
                unit_sep_ver_add=ver_address[:comma].lower()
                unit_sep_ver_add+=" "+city_state.lower()
                if(not flag):
                    if(unit_sep_ver_add==unit_sep_ad and unit==ver_unit):
                        return(True)
                elif(unit_sep_ver_add==unit_sep_ad):
                    return(True)
                #print(unit_sep_check_add,check_unit,unit_sep_ad,unit)
            return(False)
        

            
         

        

                 
            
