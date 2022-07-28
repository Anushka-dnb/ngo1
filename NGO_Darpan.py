import keyboard
from keyboard import press
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from time import sleep
from datetime import date


dte=date.today()
dte=dte.strftime("%d%m%Y")

df=pd.DataFrame()

optns=webdriver.ChromeOptions()
optns.add_argument('--ignore-certificate-errors')
optns.add_argument("--test-type")
optns.binary_location="/usr/bin/chromium"

#Download "chromedriver.exe" from"https://sites.google.com/a/chromium.org/chromedriver/downloads" 
#by getting the version from chrome--> Help --> About
#driver = webdriver.Chrome(executable_path="C:/Users/Z-UiPathIN-RPA/Downloads/chromedriver.exe")
exepath="C:/Users/DeshpandeK/OneDrive - Dun and Bradstreet/Documents/Python/chromedriver.exe"

url="https://ngodarpan.gov.in/index.php/home/statewise"
#url="https://ngodarpan.gov.in/index.php/home/statewise_ngo/222/17/3?per_page=100"
#stlist={"ARUNACHAL PRADESH", "ASSAM", "CHHATTISGARH", "DELHI", "GUJARAT", "HARYANA", "JHARKHAND", "KARNATAKA", "KERALA", "MANIPUR", "ORISSA", "PUNJAB", "RAJASTHAN", , "TELANGANA", "UTTAR PRADESH", "UTTARAKHAND", "WEST BENGAL"}
##total states=37 done=33 remaining=4
stlist={"ANDAMAN & NICOBAR ISLANDS"}
##done={"DAMAN & DIU", "ANDAMAN & NICOBAR ISLANDS", "LAKSHADWEEP", "SIKKIM", "LADAKH", "CHANDIGARH", "MIZORAM", "GOA", "TRIPURA", "PUDUCHERRY", "ARUNACHAL PRADESH", "HIMACHAL PRADESH", "NAGALAND", "JAMMU & KASHMIR", "MAHARASHTRA", "UTTARAKHAND", "ORISSA", "DELHI", "KERALA", "UTTAR PRADESH", "GUJRAT", "WEST BENGAL", "BIHAR", "MADHYA PRADESH", "TAMIL NADU", "ANDHRA PRADESH", }


try:
    for states in stlist:
        driver = webdriver.Chrome(executable_path=exepath)
        driver.get(url)
        time.sleep(1)
        classnm=driver.find_elements_by_xpath("//li[@class='bluelink11px']")
        state=driver.find_element_by_partial_link_text(states)
        state.click()
        time.sleep(3)
        view=driver.find_element_by_xpath("//select[@name='per_page'and @class='form-control']")
        view.send_keys("100")
        newurl=driver.current_url
        tables=pd.read_html(newurl)[0]
        names=tables['Name of VO/NGO']
        for name in names:
            time.sleep(3)
            driver.refresh()
            time.sleep(7)
            clckname=driver.find_element_by_partial_link_text(name) #find_element_by_link_text(name)
            clckname.click()
            time.sleep(7)
            data=driver.find_element_by_xpath("//div[@id='ngo_info_modal' and @class='modal fade in']")
            time.sleep(7)
            txt=data.text
            ##Split txt into variables and put(append) it in data table
            foundname1=txt.split("NGO DETAILS")[1]
            foundname2=foundname1.split("Unique Id")[0]
            foundname=foundname2.strip()
            num1=txt.split("Unique Id of VO/NGO")[1]
            num2=num1.split("Registration Details")[0]
            num=num2.strip()
            regwith1=txt.split("Registered With")[1]
            regwith2=regwith1.split("Type of NGO")[0]
            regwith=regwith2.strip()
            tyngo1=txt.split("Type of NGO")[1]
            tyngo2=tyngo1.split("Registration No")[0]
            tyngo=tyngo2.strip()
            regno1=txt.split("Registration No")[1]
            regno2=regno1.split("Copy of Registration")[0]
            regno=regno2.strip()
            regcert1=txt.split("Copy of Registration Certificate")[1]
            regcert2=regcert1.split("Copy of Pan")[0]
            regcert=regcert2.strip()
            pan1=txt.split("Copy of Pan Card")[1]
            pan2=pan1.split("Act name")[0]
            pan=pan2.strip()
            actname1=txt.split("Act name")[1]
            actname2=actname1.split("City of Registration")[0]
            actname=actname2.strip()
            regcity1=txt.split("City of Registration")[1]
            regcity2=regcity1.split("State of Registration")[0]
            regcity=regcity2.strip()
            regstate1=txt.split("State of Registration")[1]
            regstate2=regstate1.split("Date of Registration")[0]
            regstate=regstate2.strip()
            regdte1=txt.split("Date of Registration")[1]
            regdte2=regdte1.split("Members")[0]
            regdte=regdte2.strip()
            membs1=txt.split("Members")[1]
            membs2=membs1.split("Sector/ Key Issues")[0]
            membs=membs2.strip()
            issue1=txt.split("Key Issues")[2]
            issue2=issue1.split("Operational Area-States")[0]
            issue=issue2.strip()
            funstate1=txt.split("Operational Area-States")[1]
            funstate2=funstate1.split("Operational Area-District")[0]
            funstate=funstate2.strip()
            fundist1=txt.split("Operational Area-District")[1]
            fundist2=fundist1.split("FCRA details")[0]
            fundist=fundist2.strip()
            facra1=txt.split("FCRA details")[1]
            facra2=facra1.split("Details of Achievements")[0]
            facra=facra2.strip()
            detls1=txt.split("Details of Achievements")[1]
            detls2=detls1.split("Source of Funds")[0]
            detls=detls2.strip()
            fundsrc1=txt.split("Source of Funds")[1]
            fundsrc2=fundsrc1.split("Contact Details")[0]
            fundsrc=fundsrc2.strip()
            contactdetails=txt.split("Contact Details\n")[1]
            if "Old Address" in contactdetails:
                add1=contactdetails.split("Old Address")[1]
                add2=add1.split("City")[0]
                add=add2.strip()
            else:
                add1=txt.split("Address")[1]
                add2=add1.split("City")[0]
                add=add2.strip()
            city1=contactdetails.split("City")[1]
            city2=city1.split("State")[0]
            city=city2.strip()
            state1=contactdetails.split("State")[1]
            state2=state1.split("Telephone")[0]
            state=state2.strip()
            tel1=txt.split("Telephone")[1]
            tel2=tel1.split("Mobile No")[0]
            tel=tel2.strip()
            mob1=txt.split("Mobile No")[1]
            mob2=mob1.split("Website Url")[0]
            mob=mob2.strip()
            web1=txt.split("Website Url")[1]
            web2=web1.split("E-mail")[0]
            web=web2.strip()
            mail1=txt.split("E-mail")[1]
            mail=mail1.strip().replace("(at)","@").replace("[dot]", ".")
            tab=pd.DataFrame({"ComapnyName":foundname, "Unique Id of VO/NGO":num, "Registered With":regwith, "Type of NGO":tyngo, "Registration No":regno, "Registration Certificate":regcert, "Pan Card":pan, "Act name":actname, "City of Registration":regcity, "State of Registration":regstate, "Date of Registration":regdte, "Members":membs, "Key Issues":issue, "Operational Area-States":funstate, "Operational Area-District":fundist, "FCRA details":facra, "Details of Achievements":detls, "Source of Funds":fundsrc, "Address":add, "City":city, "State":state, "Telephone":tel, "Mobile No":mob, "Website Url":web, "E-mail":mail}, index=[0])
            df=df.append(tab)
            time.sleep(2)
            #close=driver.find_element_by_xpath("//button[@type='button' and @class='close']")# and @data-dismiss='modal and @aria-label='Close']")
            #close.click()
            driver.get(newurl)
            time.sleep(5)
            driver.refresh()
            time.sleep(5)
        #try:
        while True:
            nxt=driver.find_element_by_xpath("//a[@rel='next']").click()
            time.sleep(2)
            newpg=driver.current_url
            tables=pd.read_html(newpg)[0]
            names=tables['Name of VO/NGO']
            for name in names:
                time.sleep(3)
                driver.refresh()
                time.sleep(7)
                clckname=driver.find_element_by_link_text(name)
                clckname.click()
                time.sleep(7)
                data=driver.find_element_by_xpath("//div[@id='ngo_info_modal' and @class='modal fade in']")
                time.sleep(7)
                txt=data.text
                ##split data
                foundname1=txt.split("NGO DETAILS")[1]
                foundname2=foundname1.split("Unique Id")[0]
                foundname=foundname2.strip()
                num1=txt.split("Unique Id of VO/NGO")[1]
                num2=num1.split("Registration Details")[0]
                num=num2.strip()
                regwith1=txt.split("Registered With")[1]
                regwith2=regwith1.split("Type of NGO")[0]
                regwith=regwith2.strip()
                tyngo1=txt.split("Type of NGO")[1]
                tyngo2=tyngo1.split("Registration No")[0]
                tyngo=tyngo2.strip()
                regno1=txt.split("Registration No")[1]
                regno2=regno1.split("Copy of Registration")[0]
                regno=regno2.strip()
                regcert1=txt.split("Copy of Registration Certificate")[1]
                regcert2=regcert1.split("Copy of Pan")[0]
                regcert=regcert2.strip()
                pan1=txt.split("Copy of Pan Card")[1]
                pan2=pan1.split("Act name")[0]
                pan=pan2.strip()
                actname1=txt.split("Act name")[1]
                actname2=actname1.split("City of Registration")[0]
                actname=actname2.strip()
                regcity1=txt.split("City of Registration")[1]
                regcity2=regcity1.split("State of Registration")[0]
                regcity=regcity2.strip()
                regstate1=txt.split("State of Registration")[1]
                regstate2=regstate1.split("Date of Registration")[0]
                regstate=regstate2.strip()
                regdte1=txt.split("Date of Registration")[1]
                regdte2=regdte1.split("Members")[0]
                regdte=regdte2.strip()
                membs1=txt.split("Members")[1]
                membs2=membs1.split("Sector/ Key Issues")[0]
                membs=membs2.strip()
                issue1=txt.split("Key Issues")[2]
                issue2=issue1.split("Operational Area-States")[0]
                issue=issue2.strip()
                funstate1=txt.split("Operational Area-States")[1]
                funstate2=funstate1.split("Operational Area-District")[0]
                funstate=funstate2.strip()
                fundist1=txt.split("Operational Area-District")[1]
                fundist2=fundist1.split("FCRA details")[0]
                fundist=fundist2.strip()
                facra1=txt.split("FCRA details")[1]
                facra2=facra1.split("Details of Achievements")[0]
                facra=facra2.strip()
                detls1=txt.split("Details of Achievements")[1]
                detls2=detls1.split("Source of Funds")[0]
                detls=detls2.strip()
                fundsrc1=txt.split("Source of Funds")[1]
                fundsrc2=fundsrc1.split("Contact Details")[0]
                fundsrc=fundsrc2.strip()
                contactdetails=txt.split("Contact Details")[1]
                if "Old Address" in contactdetails:
                    add1=contactdetails.split("Old Address")[1]
                    add2=add1.split("City")[0]
                    add=add2.strip()
                else:
                    add1=txt.split("Address")[1]
                    add2=add1.split("City")[0]
                    add=add2.strip()
                city1=contactdetails.split("City")[1]
                city2=city1.split("State")[0]
                city=city2.strip()
                state1=contactdetails.split("State")[1]
                state2=state1.split("Telephone")[0]
                state=state2.strip()
                tel1=contactdetails.split("Telephone")[1]
                tel2=tel1.split("Mobile No")[0]
                tel=tel2.strip()
                mob1=contactdetails.split("Mobile No")[1]
                mob2=mob1.split("Website Url")[0]
                mob=mob2.strip()
                web1=contactdetails.split("Website Url")[1]
                web2=web1.split("E-mail")[0]
                web=web2.strip()
                mail1=contactdetails.split("E-mail")[1]
                mail=mail1.strip().replace("(at)","@").replace("[dot]", ".")
                tab=pd.DataFrame({"ComapnyName":foundname, "Unique Id of VO/NGO":num, "Registered With":regwith, "Type of NGO":tyngo, "Registration No":regno, "Registration Certificate":regcert, "Pan Card":pan, "Act name":actname, "City of Registration":regcity, "State of Registration":regstate, "Date of Registration":regdte, "Members":membs, "Key Issues":issue, "Operational Area-States":funstate, "Operational Area-District":fundist, "FCRA details":facra, "Details of Achievements":detls, "Source of Funds":fundsrc, "Address":add, "City":city, "State":state, "Telephone":tel, "Mobile No":mob, "Website Url":web, "E-mail":mail}, index=[0])
                df=df.append(tab)
                time.sleep(2)
                driver.get(newpg)
                time.sleep(5)
                driver.refresh()
                time.sleep(5)
        #except:
            #driver.close()
            #print("no new page")
            #continue
except:
    pass
#driver.close()

wb=pd.ExcelWriter("C:/Users/DeshpandeK/OneDrive - Dun and Bradstreet/Documents/Apeda/NGODarpan/NGO_DARPAN_ANDAMAN NICOBAR ISLAND_July2022_output"+dte+".xlsx")
df.to_excel(wb, sheet_name="Sheet1", index=False)
wb.save()
            
            