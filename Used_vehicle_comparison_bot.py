import ClointFusion as cf
from tkinter.constants import TRUE
import time
from datetime import date
from zipfile import ZipFile
import os
from os.path import basename

#openiing mahindra first choice website
def open_website_mahindra():

    browser_state = False
    
    try:
        browser_state = cf.launch_website_h(Mahindra_first_link)
        
    except:
        print("Error in opening Mahindra First choice website")

    finally:
        return browser_state

#selecting location in mahindra first choice website
def location_select_mahindra():
    try:
        cf.browser_write_h(city_name,User_Visible_Text_Element="Select Your City")
        time.sleep(0.5)
        cf.browser_mouse_click_h(city_name)

    except:
        print("location cannot be selected")

#selecting car in mahindra first choice website
def car_select_mahindra():
    
    try:
    
        cf.browser_write_h(car_name,User_Visible_Text_Element="Select Brand")
        time.sleep(0.5)
        cf.browser_mouse_click_h("FIND YOUR CAR")
        time.sleep(0.5)


    except:
        print("car cannot be selected")


#creating mahindrafirst choice spreadsheet
def create_excel_sheet_mahindra():
    try:
        cf.excel_create_excel_file_in_given_folder(folder_location[:-1],excelFileName=sheet_name_mahindra )
    
    except:
        print("excel sheet cannot be created")  

#storing the information in the mahindrafirstchoice spreadsheet 
def store_excel_sheet_mahindra():  

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='buyer_header']//h3[@class='buyer_make Stock_viewed'] ")
        for len in d:
            g=str(len).split(">")
            t=g[1].split("<")
            print(t[0])
            cf.excel_set_single_cell(sheet_location_mahindra,columnName="Model name",cellNumber=i,setText=t[0])
            i=i+1

    except:
        print("Error in collecting data")
    try:

        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='buyer_header']//h3[@class='buyer_variant']")
        i=0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
            print(t[0])
            cf.excel_set_single_cell(sheet_location_mahindra,columnName="Car name",cellNumber=i,setText=t[0])
            i=i+1
    except:
        print("Error in collecting data")

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='stock_card_caption']//span[@class='car_price']")
        i=0
        for price in d:
            print(price.web_element.text)
            cf.excel_set_single_cell(sheet_location_mahindra,columnName="Car price",cellNumber=i,setText=price.web_element.text)
            i=i+1


    except:
        print("Error in collecting data")

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='stock_card_caption']//span[@class='overview_name']")
        d
        i=0
        q=0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
  
            if(i%4==0):
                print(t[0])
                cf.excel_set_single_cell(sheet_location_mahindra,columnName="Kilometres used",cellNumber=q,setText=t[0])
                q=q+1
            i=i+1
    except:
        print("Error collecting info")

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='stock_card_caption']//span[@class='overview_name']")
        d
        i=2
        q=0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
  
            if(i%4==0):
                if t[0] == "":
                  print("-")
                  cf.excel_set_single_cell(sheet_location_mahindra,columnName="Car type",cellNumber=q,setText="-")
                  q=q+1
                else:
                  print(t[0])
                  cf.excel_set_single_cell(sheet_location_mahindra,columnName="Car type",cellNumber=q,setText=t[0])
                  q=q+1
            i=i+1

    except:
        print("Error collecting info")

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='stock_card_caption']//span[@class='overview_name']")
        q=0
        i=3
        for len in d:
            g=str(len).split(">")
            t=g[1].split("<")
            if(i%4==0):
                print(t[0])
                cf.excel_set_single_cell(sheet_location_mahindra,columnName="Engine type",cellNumber=q,setText=t[0])
                q=q+1
            i=i+1

    except:
        print("Error collecting info")

    try:
        d =cf.browser_locate_elements_h("//*[@id='buyer_stock_results']//div[@class='stock_card_caption']//span[@class='overview_name']")
        i=5
        q=0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
  
            if(i%4==0):
                print(t[0])
                cf.excel_set_single_cell(sheet_location_mahindra,columnName="Owner",cellNumber=q,setText=t[0])
                q=q+1
            i=i+1

    except:
        print("Error in collecting info")


#opening the cars24 website
def open_website():

    browser_state = False
    
    try:
        browser_state = cf.launch_website_h(cars24_link)
        
    except:
        print("Error in opening cars24 website")

    finally:
        return browser_state

#selecting the location in cars24 website
def location_select():
    try:
        cf.browser_mouse_click_h("SELECT MANUALLY")
        time.sleep(0.5)
        cf.browser_write_h(city_name,User_Visible_Text_Element="Search City")
        time.sleep(0.5)
        cf.browser_mouse_click_h(city_name)

    except:
        print("location cannot be selected")


#selecting the car in cars24 website
def car_select():   
    try:
        cf.browser_mouse_click_h("VIEW ALL CARS")
        time.sleep(0.5)
        cf.browser_mouse_click_h("Find your dream car with us")
        time.sleep(0.5)
        cf.browser_write_h(car_name,User_Visible_Text_Element="Find your dream car with us")
        time.sleep(0.5)
        cf.key_write_enter(strMsg=" ")
        time.sleep(0.5)


    except:
        print("car cannot be selected")

#creating the cars24 spreadsheet
def create_excel_sheet():
    
    try:
        cf.excel_create_excel_file_in_given_folder(folder_location[:-1],excelFileName=sheet_name )
    
    except:
        print("excel sheet cannot be created")


#storing the required details in cars24 spreadsheet
def store_excel_sheet():
    
    time.sleep(1)
    
    try:
        d =cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//h2[@itemprop='name']")
        i= 0
        for len in d:
            g=str(len).split(">")
            h=g[1].split("<")
            print(h[0])
            cf.excel_set_single_cell(sheet_location,columnName="Name",cellNumber=i,setText=h[0])
            i=i+1
  
    except:
        print("error in collecting the car names") 
    time.sleep(1)
   
    try:
        c=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//h3")
        i=0
        for len in c:
            g=str(len).split(">")
            h=g[1].split("<")
            print(h[0])
            cf.excel_set_single_cell(sheet_location,columnName="Price",cellNumber=i,setText=h[0])
            i=i+1

    except:
        print("error in collecting the car price")

  
    time.sleep(1)


    try:
        c=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//p//span")

        i=0
        q=0
        for len in c:
            g=str(len).split(">")
            h=g[1].split("<")
            t=g[1].split("<")
  
            if(i%4==0):
                print(t[0]) 
                cf.excel_set_single_cell(sheet_location,columnName="Kilometres used",cellNumber=q,setText=t[0])
                q=q+1 
            i=i+1
            
    except:
        print("error in collecting the kilometres used")

  
    time.sleep(1)


    try:
        g=cf.browser_locate_elements_h("//div[@itemprop='itemOffered']//p//span[@itemprop='name']")
        i=0
        for len in g:
          g=str(len).split(">")
          h=g[1].split("<")
          print(h[0])
          cf.excel_set_single_cell(sheet_location,columnName="Engine type",cellNumber=i,setText=h[0])
          i=i+1
   
    except:
        print("error in collecting the engine type ")

def get_all_file_paths(directory):

 # initializing empty file paths list
 file_paths = []

 # crawling through directory and subdirectories
 for root, directories,files in os.walk(directory):
  for filename in files:
   # join the two strings in order to form the full filepath.
   filepath = os.path.join(root, filename)
   file_paths.append(filepath)

 # returning all file paths
 return file_paths

def zip_files():
 # path to folder which needs to be zipped
 directory = 'C:\Cars24_car_details_download_automation'

 # calling function to get all file paths in the directory
 file_paths = get_all_file_paths(directory)

 # printing the list of all files to be zipped
 print('Following files will be zipped in this program:')
 for file_name in file_paths:
    if "xlsx" in str(file_name): 
        print(file_name)

 # writing files to a zipfile
 with ZipFile('Cars_report.zip','w') as zip:
  # writing each file one by one
  for file in file_paths:
    if "xlsx" in str(file):
        zip.write(file)

 print('All files zipped successfully!')


#Outlook email function
def send_outlook_email():
    try:

        #getting the outlook credentials from json file
        outlook_details = cf.file_get_json_details(path_of_json_file=CREDENTIALS_JSON,section='Outlook')

        outlook_username = outlook_details.get('username')
        outlook_password = outlook_details.get('password')
        to = outlook_details.get('send_to')

        cf.browser_navigate_h('outlook.com')
        time.sleep(1)
        cf.browser_mouse_click_h('Sign in')
        time.sleep(0.5)

        cf.browser_write_h(outlook_username,User_Visible_Text_Element='Email, phone, or Skype')
        time.sleep(0.5)
        cf.browser_mouse_click_h('Next')

        time.sleep(0.5)

        cf.browser_write_h(outlook_password,User_Visible_Text_Element='Password')
        time.sleep(1)
        cf.browser_mouse_click_h('Sign in')
        time.sleep(1)

        cf.browser_mouse_click_h('New message')

        time.sleep(1)
        cf.browser_write_h(to,User_Visible_Text_Element='To')
        time.sleep(1)

        cf.browser_write_h('car details from cars 24',User_Visible_Text_Element='Add a subject')
        
        body_elem = cf.browser_locate_element_h("//*[@aria-label='Message body']")
        cf.browser_write_h('Please find the attached Report.\n\n\nThanks & Regards\nMohammad Yawar',User_Visible_Text_Element=body_elem)
        
        time.sleep(1)

        cf.browser_mouse_click_h(User_Visible_Text_Element='Attach')
        time.sleep(0.5)
        cf.browser_mouse_click_h(User_Visible_Text_Element='Browse this computer')
        time.sleep(0.5)

        #sending the zip file
        cf.key_write_enter(strMsg='C:\Cars24_car_details_download_automation\Cars_report.zip')
        time.sleep(1)

        cf.browser_mouse_click_h('Send')


    except:
        print("Error in Sending Outlook Email")


#calling all the necessary functions needed for cars24 website   
def cars_24():
    try:    

        browser_state= open_website()
        time.sleep(1)
     
     
        if browser_state==TRUE:
   
            #setting the location in cars24 website
            location_select()
            time.sleep(1)

            #selecting the car model
            car_select()
            time.sleep(1)
            
            #creating the excel sheet
            create_excel_sheet()
            time.sleep(1)
            
            #entering the details in excel sheet
            store_excel_sheet()
            time.sleep(1)
        
        else:
            print("browser not opened")
    
    except:
        print("error")


#calling all the necessary functions needed for mahindrafirst website
def mahindra_first():
    try:
        browser_state= open_website_mahindra()
        time.sleep(1)

        if browser_state==TRUE:
   
            #setting the location in mahindra_first_choice website
            location_select_mahindra()
            time.sleep(1)

            #selecting the car model
            car_select_mahindra()
            time.sleep(1)

            #creating the excel sheet
            create_excel_sheet_mahindra()
            time.sleep(1)

            #storing in excel sheet
            store_excel_sheet_mahindra()
        else:
            print("browser not opened")
    except:
        print("error")

# Zip the files from given directory that matches the filter



if __name__ == '__main__':

    #assigning all the variables
    date_today = str(date.today())
    cars24_link = "https://www.cars24.com"
    Mahindra_first_link = "https://www.mahindrafirstchoice.com/"

    #getting Outlook credential details from json file
    CREDENTIALS_JSON = "C:\Cars24_car_details_download_automation\credentials.json"

    #getting car and city information from json
    DETAILS_JSON = "C:\Cars24_car_details_download_automation\details.json"

    #storing the city names in a list
    city_details = cf.file_get_json_details(path_of_json_file=DETAILS_JSON,section='city')
    city_names=[]
    city_names.append(city_details.get('city_name_1'))
    city_names.append(city_details.get('city_name_2'))
    city_names.append(city_details.get('city_name_3'))
    city_names.append(city_details.get('city_name_4'))
    city_names.append(city_details.get('city_name_5'))

    #soring the car name
    DETAILS_JSON = "C:\Cars24_car_details_download_automation\details.json"
    car_details = cf.file_get_json_details(path_of_json_file=DETAILS_JSON,section='car_details')
    car_name=car_details.get('car_name')

    #iterating for all the cities
    for city_name in city_names:

        #setting the sheet name according to the city,car and date
        sheet_name = car_name +"_"+ city_name + "_" + date_today + "_cars24"
        sheet_name_mahindra = car_name +"_"+ city_name + "_" + date_today + "_MahindraFirstChoice"
        folder_location='C:\Cars24_car_details_download_automation\ '
        sheet_location= folder_location[:-1] + sheet_name + ".xlsx" 
        sheet_location_mahindra= folder_location[:-1] + sheet_name_mahindra + ".xlsx"

        #calling the cars24 website function
        cars_24()

        #calling the mahindra_first_choice website function
        mahindra_first()

    #zipping all the xlsx files  
    zip_files()
    #sending the outlook email
    send_outlook_email()
    time.sleep(3)
    cf.browser_quit_h()
