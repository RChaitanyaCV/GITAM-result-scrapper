import xlsxwriter as xw
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
wb=xw.Workbook('GITAM_RESULTS.xlsx')
ws=wb.add_worksheet()
ws.write(0,0,"Pin Number")
ws.write(0,1,"Name")
ws.write(0,2,"Sem 1")
ws.write(0,3,"Sem 2")
ws.write(0,4,"Sem 3")
ws.write(0,5,"Sem 4")
ws.write(0,6,"Sem 5")
ws.write(0,7,"Sem 6")
ws.write(0,8,"Sem 7")
ws.write(0,9,"Sem 8")
t=int(input('Enter your pin number\n'))
t=int(t/100)
t=t*100
driver=webdriver.Chrome(executable_path='C:/Users/pathto/chromedriver.exe') # Insert chromedriver path here
driver.set_page_load_timeout("5")
driver.get("https://doeresults.gitam.edu/onlineresults/pages/newgrdcrdinput1.aspx") # Result page link
for i in range(1,67): # Iteration for all pin numbers
    t+=1
    for j in range(1,9): # Iterator for all sems
        driver.find_element_by_name("cbosem").send_keys(str(j))
        driver.find_element_by_name("txtreg").send_keys(Keys.BACK_SPACE)
        driver.find_element_by_name("txtreg").send_keys(str(t))
        driver.find_element_by_name("Button1").send_keys(Keys.ENTER)
        try:
            if j==1:
                p=driver.find_element_by_id("lblregdno")
                p_t=p.text
                ws.write(i,0,format(p_t))
                n=driver.find_element_by_id("lblname")
                n_t=n.text
                ws.write(i,1,format(n_t))
            s=driver.find_element_by_id("lblgpa")
            s_t=s.text
            ws.write(i,j+1,format(s_t))
        except:
            print(str(t)+"\n")
        time.sleep(1)
        driver.get("https://doeresults.gitam.edu/onlineresults/pages/Newgrdcrdinput1.aspx")
        time.sleep(1)
driver.quit()
wb.close()