# Spokeo Bot

## Description
- A bot that takes a name(or names) and its corresponding address(or addresses) from an excel file as input, finds a matching profile on spokeo.com and if found, and saves its contact information as output.
- It was made to eliminate the need to do the above task for large number of contacts manually.
- The main technologies and modules used were **python**, **Selenium**, **undetected-chrome-driver** and **openpyxl**.
 
## How to install and use
Currently this application is only supported on Windows OS(7 and above).<br>
1. Install Python 3.x including pip.<br>Guide- https://www.digitalocean.com/community/tutorials/install-python-windows-10
2. Run the following commands on cmd.
    
       py -m pip --upgrade pip
       py -m pip install openpyxl ordered-set selenium undetected-chromedriver
3. Install Chrome Browser (Don't install version 103.x) and enter version number in find_info/find_info.py
4. Create a chrome profile and login to spokeo.
5. Enter the profile path and profile name in the contructor of find_info in run.py.
6. execute run.py.

## Credits
- ultrafunkhamsterdam for undetected chrome driver<br>https://github.com/ultrafunkamsterdam/undetected-chromedriver
- Jim from JimShapedCoding and freecodecamp.org<br>https://www.youtube.com/channel/UCU8d7rcShA7MGuDyYH1aWGg<br>freecodecamp.org

