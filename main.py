#Tkinter
import tkinter as tk
from tkinter import BitmapImage
from tkinter import messagebox
import json
from datetime import date
#SELENIUM
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from selenium.common.exceptions import TimeoutException
from datetime import datetime
from selenium.common.exceptions import TimeoutException, NoSuchElementException
#PANDA FOR READ CSV/XLSX FILES
import pandas as pd
import os
#OTHER IMPORTS
from datetime import datetime
import datetime
from datetime import date
import requests
from time import sleep

##ROOT AND INTERFACE
root=tk.Tk()
root.title('Automatizacion con Selenium y Python')
root.configure(background='#02b5dd')

#variables
Url_var=tk.StringVar()
Email_var=tk.StringVar()
Password_var=tk.StringVar()
ConsoleOutput=tk.StringVar()
StartFrom = tk.IntVar()
#arrays


arrayForFormCopy=[
    #0
    "nombre",
    #1
    "apellidos",
    #2
    "email",
    #3
    "dni",
    #4
    "fecha_nacimiento",
    #5
    "genero"]

arrayForXpathCopy = [
    #0
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[4]/div[1]/div/input",
    #1
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[4]/div[2]/div/input",
    #2
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[5]/div[1]/div/input",
    #3
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[5]/div[2]/div/input",
    #4
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[5]/div[3]/input",
    #5
    [
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[6]/div/div/div[1]/div/label",
    "/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[6]/div/div/div[2]/div/label"
    ],

]

#nombre adulto
# /html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[1]/div/input

#apellidos adulto
# /html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[2]/div/input

#DNI ADULTO 
# /html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[3]/div/input
arrayForXpathOptionalCopy=[
    #0
    [ "movil","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[10]/div[2]/div/input"],
    #1
    [ "movil_urgencias","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[10]/div[3]/div/input"],
    #2
    [ "prefijo_telefono","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[10]/div[1]/div[1]/input"],
    #3
    [ "email_confirm","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[8]/div/div/input"],
    #4
    [ "pais_origen","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[8]/div/input"],
    #5 group selection with #6 #7
    [ "pais","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[4]/div/input"],
    #6
    [ "provincia","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[5]/div/div[1]/input"],
    #7
    [ "ciudad","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[6]/div/input"],
    #8
    [ "direccion","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[1]/div/input"],
    #9
    [ "codigo_postal","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[9]/div[7]/div/input"],
    #10 #permiso Adulto
    ["permiso_paterno", {"nombre_adulto":"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[1]/div/input","apellidos_adulto":"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[2]/div/input","Dni_Adulto":"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[12]/div[3]/div/input"}],
    #11
    [ "participante_local","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[11]/div[2]/div/label"],
    #12
    [ "participante_federado","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[11]/div[2]/div/label"],
    #13
    [ "foto_carnet","/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[13]/div/input"],
]

source_text_excel = []
errorArray=[]
#CopyArray
arrayForForm = arrayForFormCopy.copy()
arrayForXpath = arrayForXpathCopy.copy()
arrayForXpathOptional = arrayForXpathOptionalCopy.copy()
#countRows
lenexcel = pd.read_excel('test1.xlsx', engine='openpyxl',dtype=str)
lenexcelcount = lenexcel.count()
lenexcelcount = len(lenexcel)

lenexcellen = len(lenexcel)
lenexcelshape = lenexcel.shape[0]
print(lenexcelcount)
print(lenexcellen)
lenexcel=lenexcellen-3

def calculateAge(dateToCalculate):
    birth = datetime.datetime.strptime(dateToCalculate, "%d/%m/%Y")
    today = date.today()
    return today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))


def SeleniumWebDrive(url):  
    global arrayForForm
    global arrayForXpath
    global arrayForXpathOptional

    arrayForForm = arrayForFormCopy.copy()
    arrayForXpath = arrayForXpathCopy.copy()
    arrayForXpathOptional = arrayForXpathOptionalCopy.copy()
    urlafterinput = url
    print(urlafterinput)
    driver = webdriver.Chrome(executable_path=r"drivers\chromedriver.exe")
    driver.set_window_size(1,1)
    driver.get(baseUrl)
    if TestFor404(driver,baseUrl,"/") is False:
        return print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
    driver.set_window_position(0, 0)
    driver.set_window_size(1024, 768)
    loginAndSubmit(driver)
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, "SidebarHome")))
    driver.implicitly_wait(4)
    driver.get(url+"#inscriptionconfig")
    sleep(1)
    if TestFor404(driver,url,"#inscriptionconfig") is False:
       return print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
    print(driver.current_url)
    closeCokies = driver.find_element_by_id('rcc-confirm-button')
    closeCokies.click()
    
    getcheckedAtributes(driver)
    driver.implicitly_wait(4)  
    fromStart = 0
    fromStart= StartFrom.get()-1
    if fromStart == -1:
        fromStart = 0
    
    for i in range(fromStart,lenexcel):
        age=0
        arrayForXpath[5] = [
        str("/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[6]/div/div/div[1]/div/label"),
        str("/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[6]/div/div/div[2]/div/label")]
        handleCustomInsert("Start----------- " + str(i+1) + " -----------Test",text)
        if TestFor404(driver,url,"#inscriptionconfig") is False:
            print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
            break
        print(driver.current_url)
        global source_text_excel
        if readExcelFile(i) is False:
            break
        driver.get(url +"/inscription/create")
        sleep(1)  
        if TestFor404(driver,url,"/inscription/create") is False:
            print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
            break      
        if checkIfElementExist(driver,"/html/body/div/div[3]/div/div/div[2]/div[2]/div[3]/button[2]") is False:
            driver.get(url +"/inscription/create")
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/div/div/div[2]/div[2]/h2")))
        switchtoES = driver.find_element_by_xpath("/html/body/div/div[3]/div/div/div[1]/div/div[1]/div/div[2]/button[1]")
        switchtoES.click()
        if TestFor404(driver,url,"/inscription/create") is False:
            print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
            break 
        WriteForm(driver,url)
        ok = False
        while not ok:
            sleep(0.5)
            trysubmit =driver.find_element_by_xpath("/html/body/div/div[3]/div/div/div[2]/div[2]/div[3]/button[2]")
            trysubmit.click()
            WebDriverWait(driver, 5)
            if checkIfElementExist(driver,"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[3]/h3"):
                #TRUE CHECK FOR ERRORS
               global errorArray
               errorArray=[]
               checkErrorsAfterElementWrited(driver)
               handleErrorInsert(text,i)
            elif checkIfElementExist(driver,"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/div/div/div/div[1]/div/div/div/label"):
                print("nie")
            else:
                handleCustomInsert("El test ha tenido exito ",text)
                checkIfElementExist(driver,"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/div/div/div/div[1]/div/div/div/label")
                trysubmitInscription =driver.find_element_by_xpath("/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div/button")
                trysubmitInscription.click()
                WebDriverWait(driver, 5)
                sleep(2)
                #to do try to get wich error have next  
            ok = True
            if TestFor404(driver,url,"#inscriptions") is False:
                print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
                break 
        print(errorArray)
        handleCustomInsert("\n",text)
        if TestFor404(driver,url,"#inscriptionconfig") is False:
            print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
            break
    driver.close()


def appendToArray(valueXpath,valueName,arrayToName,arrayToXpath):
    if valueName == "permiso_paterno":
       dictStructure = valueXpath
       for key,values in dictStructure.items():
           arrayToName.append(str(key))
           arrayToXpath.append(str(values))
    else:
        arrayToName.append(valueName)
        arrayToXpath.append(valueXpath)

def readExcelFile(row):
    try:
        global source_text_excel
        source_text_excel = []
        source_text_excel = pd.read_excel('test1.xlsx', engine='openpyxl',dtype=str)
        source_text_excel = source_text_excel.loc[row]
        source_text_excel = source_text_excel.loc[arrayForForm]
        print(source_text_excel)
        if source_text_excel.isnull().values.any() == True:
            raise ValueError
        return True
    except ValueError:
        return False 

#open errors.json file 
with open('errors.json') as f:
    data = json.load(f)

#Call Error JSON 
def errorhandler(error):
    testerr = data.get(str(error))
    if testerr == None:
        messagebox.showinfo("Error","Escribe a bogdan " + str(error))
    else:
        messagebox.showinfo("Error", testerr)

def submit():
    test = Url_var.get() 
    seleniumStart(test)

def getdatatoput():
    datatest = "bogdan"
    return datatest

def seleniumStart(url):
   res = requests.get(url)
   print(res.status_code)
   if res.status_code != 200:
       errorhandler(res.status_code)
   else:
       SeleniumWebDrive(url)

def SelectAndClick_by_id(driver,componentName,elemId,sendkey):
    try:
        componentName = driver.find_element_by_id(elemId)
        componentName.click()
        if sendkey != "submit":
            componentName.send_keys(sendkey)
    except NoSuchElementException:
        print("SelectAndClick_by_id")

def SelectClickSubmit_by_xpath(driver,componentName,elemXpath,sendkey):
    try:
        componentName = driver.find_element_by_xpath(elemXpath)
        componentName.send_keys('\n') 
        componentName.send_keys(sendkey)
        sleep(0.5)
    except NoSuchElementException:
        print("SelectClickSubmit_by_xpath")

def Click_by_xpath(driver,componentName,elemXpath):
    try:
        componentName = driver.find_element_by_xpath(elemXpath)
        componentName.click()
        check = driver.find_element_by_id("root")
        check = check.click()
    except NoSuchElementException:
        print("Click_by_xpath")


def UploadFile(driver,componentName,elemXpath):
    try:
        fotopath = os.path.abspath("fotos/foto.png")
        componentName = driver.find_element_by_xpath(elemXpath)
        componentName.send_keys(fotopath)
    except NoSuchElementException:
        print("ok") 


#Check Errors
def checkErrorsAfterElementWrited(driver):
    rgb2 = "rgba(159, 58, 56, 1)"
    age = calculateAge(str(source_text_excel["fecha_nacimiento"]))
    for i in range(0,len(arrayForForm)):
        if arrayForForm[i] == "foto_carnet" and checkIfElementExist(driver,"/html/body/div/div[3]/div/div/div[2]/div[2]/div[1]/div[13]/div/div/div"):
            errorArray.append(arrayForForm[i])
            break
        elif (arrayForForm[i] == "nombre_adulto" or arrayForForm[i]== "apellidos_adulto" or arrayForForm[i]=="Dni_Adulto") and age> 18:
            continue
        testvalor= driver.find_element_by_xpath(arrayForXpath[i])
        rgb1 = testvalor.value_of_css_property("color")
        if str(rgb1) == str(rgb2):
            errorArray.append(arrayForForm[i])

def deleteIfMajor(nameindex):
    deletehere = arrayForForm.index(nameindex)
    arrayForForm.pop(deletehere)
    arrayForXpath.pop(deletehere)
    source_text_excel.drop(nameindex,inplace=True)


def clickEmptySpace(driver):
    check = driver.find_element_by_id("root")
    check = check.click()
    sleep(0.5)
#Selenium auto  
def WriteForm(driver,url):
    if TestFor404(driver,url,"/inscription/create") is False:
        return print("el servidor tiene problemas por favor intenta mas tarde o ponte en contancto con un desarollador")
    i=-1
    j=0
    age = calculateAge(str(source_text_excel["fecha_nacimiento"]))
    while j < len(arrayForForm):
        i=i+1
        if i == 5: 
            if source_text_excel[i] == "Masculino":
                Click_by_xpath(driver,str(arrayForForm[i]),arrayForXpath[i][0])
                arrayForXpath[i] = arrayForXpath[i][0]
            else:
                Click_by_xpath(driver,arrayForForm[i],arrayForXpath[i][1])
                arrayForXpath[i] = arrayForXpath[i][1]   
        elif arrayForForm[i]== "participante_local":
            if source_text_excel[i] == "Si":
                Click_by_xpath(driver,arrayForForm[i],arrayForXpath[i])
        elif arrayForForm[i]== "participante_federado":
            if source_text_excel[i] =="Si":
                Click_by_xpath(driver,arrayForForm[i],arrayForXpath[i])
        elif arrayForForm[i] == "foto_carnet":
            UploadFile(driver,arrayForForm[i],arrayForXpath[i])
        elif (arrayForForm[i] == "nombre_adulto" or arrayForForm[i]== "apellidos_adulto" or arrayForForm[i]=="Dni_Adulto") and age> 18:
            j=j+1
            continue
        else:
            SelectClickSubmit_by_xpath(driver,arrayForForm[i],str(arrayForXpath[i]),str(source_text_excel[i]))
        j = j + 1
    

def checkIfElementExist(driver,value):
    try:
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, value)))
        return True
    except TimeoutException:
        return False


def AutofillDNIDuplicated(driver):
    pass

def TestFor404(driver,url,wheretogo):
    #check if error is 404
    if driver.current_url == baseUrl + "404":
        #try to refresh 3 times
        k=0
        while k < 3:
            driver.get(url + wheretogo)
            if driver.current_url == baseUrl + "404":
                k=k+1
                if k>3:
                    return False
            else:
                return True 
    else:
        return True



baseUrl="http://ec2-15-188-81-205.eu-west-3.compute.amazonaws.com:3000/"



def loginAndSubmit(driver):
    email = Email_var.get()
    password= Password_var.get()
    SelectAndClick_by_id(driver,"email","LoginEmail",str(email))
    SelectAndClick_by_id(driver,"password","LoginPassword",str(password))
    SelectAndClick_by_id(driver,"submit","LoginButton","submit")
#Get Atributes --------- 2 ---------

def getcheckedAtributes(driver):
    print("GetcheckedAtributes-START")
    i=1
    j=1
    n=-1
    location=0
    for i in range(1,5):
        for j in range(1,5):
            n = n +1
            if j==3 and i == 4:
                return print("GetcheckedAtributes-DONE")
            elif i==2 and j==2:
                RowAtribute = "/html/body/div/div[3]/div/div/div[2]/div[3]/div[5]/div[" + str(i) + "]/div["+ str(j) +"]/div/input"
                elemprintt=driver.find_element_by_xpath(RowAtribute).is_selected()
                if elemprintt == True:
                    location = 1
                    appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
            elif i==2 and j==3:
                RowAtribute = "/html/body/div/div[3]/div/div/div[2]/div[3]/div[5]/div[" + str(i) + "]/div["+ str(j) +"]/div/input"
                elemprintt=driver.find_element_by_xpath(RowAtribute).is_selected()
                if elemprintt == True:
                    if location == 1:
                        appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
                        location = 2 
                    else:
                        appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
                        appendToArray(arrayForXpathOptional[n-1][1],str(arrayForXpathOptional[n-1][0]),arrayForForm,arrayForXpath)
                        location= 2
            elif i==2 and j == 4:
                RowAtribute = "/html/body/div/div[3]/div/div/div[2]/div[3]/div[5]/div[" + str(i) + "]/div["+ str(j) +"]/div/input"
                elemprintt=driver.find_element_by_xpath(RowAtribute).is_selected()
                if elemprintt == True:
                    if location == 1:
                        appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
                        appendToArray(arrayForXpathOptional[n-1][1],str(arrayForXpathOptional[n-1][0]),arrayForForm,arrayForXpath)
                    elif location == 2:
                        appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
                    else:
                        appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)
                        appendToArray(arrayForXpathOptional[n-1][1],str(arrayForXpathOptional[n-1][0]),arrayForForm,arrayForXpath)
                        appendToArray(arrayForXpathOptional[n-2][1],str(arrayForXpathOptional[n-2][0]),arrayForForm,arrayForXpath)
            elif i==3 and j==3:
                RowAtribute = "/html/body/div/div[3]/div/div/div[2]/div[3]/div[5]/div[" + str(i) + "]/div["+ str(j) +"]/div/input"
                elemprintt=driver.find_element_by_xpath(RowAtribute).is_selected()
                if elemprintt == True:
                        appendToArray(arrayForXpathOptional[n][1],arrayForXpathOptional[n][0],arrayForForm,arrayForXpath)
            else:
                RowAtribute = "/html/body/div/div[3]/div/div/div[2]/div[3]/div[5]/div[" + str(i) + "]/div["+ str(j) +"]/div/input"
                elemprintt=driver.find_element_by_xpath(RowAtribute).is_selected()
                if elemprintt == True:
                    appendToArray(arrayForXpathOptional[n][1],str(arrayForXpathOptional[n][0]),arrayForForm,arrayForXpath)

def checkLoginAndStoreToVariableCredentials():
    email = Email_var.get()
    password= Password_var.get()
    header={"Content-Type":"application/json"}
    loadcredentials = {"email":email,"password":password}
    try:
        api = "http://ec2-15-188-81-205.eu-west-3.compute.amazonaws.com:10010/auth"
        res = requests.post(api, data =json.dumps(loadcredentials),headers=header )
        res = res.json()
        if res["token"]:
            #to put grid
            url_label.grid(row=0,column=0,padx=10, pady=10)
            url_entry.grid(row=0,column=1,padx=10, pady=10)
            btn_submit_URL.grid(row=0,column=2,padx=10, pady=10)
    except ValueError:
        print("try again")  


#move up
def handleErrorInsert(where,nrtest):
    joinedarray = ",".join(errorArray)
    modeltext = "La modalidad ha tenido error con test nr " + str(nrtest+1) + " en los campos : " + joinedarray   
    where.insert(tk.END, modeltext+" \n")

def handleCustomInsert(textVal,where):
    where.insert(tk.END, textVal+" \n")
def gridremover():
    madeByBar.grid_remove()

frameConsole = tk.Frame(root, bg="#F1F2EB")
frameControls = tk.Frame(root,bg="#02b5dd")

#URL SUBMIT
url_label = tk.Label(frameControls,text = " URL TO TEST ",bg="#02b5dd")
url_entry = tk.Entry(frameControls,textvariable= Url_var,bg="#F1F2EB",relief='flat',width = 25)

Start_test_from_Label = tk.Label(frameControls,text = "Start from ",bg="#02b5dd" )
Start_test_from_Entry = tk.Entry(frameControls,textvariable= StartFrom,bg="#F1F2EB",relief='flat',width = 25 )

emailLoginForm_label = tk.Label(frameControls,text = " Email ",bg="#02b5dd")
emailLoginForm_entry = tk.Entry(frameControls,textvariable= Email_var,bg="#F1F2EB",relief='flat',width = 25)

passwordLoginForm_label = tk.Label(frameControls,text = " Password ", bg="#02b5dd")
passwordLoginForm_entry = tk.Entry(frameControls,textvariable=Password_var, show='*',bg="#F1F2EB",relief='flat',width = 25)
textVal="bogdan"


frameConsole['borderwidth'] = 0
frameConsole['relief'] = 'flat'

btn_submit_Login = tk.Button(frameControls, text="Login",command = checkLoginAndStoreToVariableCredentials,bg="#F1F2EB", relief='flat',width = 10)

#BUTTON SUBMIT URL

btn_submit_URL = tk.Button(frameControls, text="Empezar",command = submit,bg="#F1F2EB",relief='flat')

#madeBy
madeByBar = tk.Label(frameControls, text = "@Soluciones de Gestión Nortic",bg="#02b5dd",anchor="w",fg='#ffffff')
#madeBy 

#console
labelConsole = tk.Label(frameConsole,text=ConsoleOutput,bg="#F1F2EB",)
textStatusTesting = tk.Text(frameConsole,height = 4,bg="#F1F2EB")
textStatusTesting.pack(side=tk.TOP,fill=tk.Y, expand=False)
textStatusTesting.insert(tk.END, "Testing....")

text = tk.Text(frameConsole,bg="#F1F2EB")

text.pack(side=tk.TOP, fill=tk.Y)

# Options - CheckBox

#GRID


#URL GRID
madeByBar.grid(row=7 ,column=2, columnspan=2 )

frameControls.grid(row=0,column=0 ,rowspan=2)
frameConsole.grid(row=0,column=1)
#LOGIN GRID

emailLoginForm_label.grid(row=2,column=0 ,padx=10, pady=10)
emailLoginForm_entry.grid(row=2,column=1 ,padx=10, pady=10)

passwordLoginForm_label.grid(row=3,column=0,padx=10, pady=10)
passwordLoginForm_entry.grid(row=3,column=1,padx=10, pady=10)
Start_test_from_Label.grid(row=4,column=0)
Start_test_from_Entry.grid(row=4,column=1)
btn_submit_Login.grid (row =5, column=1,padx=10, pady=10 )


root.mainloop()