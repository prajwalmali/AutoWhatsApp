'''

Ultimate WhatsApp Automation

Repo Owner :- PrajwalCyberGod

https://www.linkedin.com/in/prajwalmali

'''

# Features

'''

Automatically Sends Same Or Different Dynamic Messages Or Attachments Or Both To List Of Peoples Or Groups Or Spam 
Detailed Documentation on GitHub

'''

# Driver Setup

'''

(You Need To Download Your Browsers Driver From Site Given Below) (*** IMPORTANT *** - Check Version Carefully)

(Link For Chrome Driver - https://chromedriver.chromium.org/downloads)

(Link For Firefox Driver - https://github.com/mozilla/geckodriver/releases)

(Link For Edge Driver - https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)

(Link For Safari Driver - https://webkit.org/blog/6900/webdriver-support-in-safari-10/)

'''

# Requirements

'''

Install By Pip If Not Done Yet
Paste Below Commands In Command Prompt For Windows :-
(For Linux, Do I Need To Tell You ?) 

pip install -r requirements.txt 

OR

pip install selenium 
pip install PyAutoIt
pip install pandas
pip install pyperclip
pip install openpyxl 
pip install tqdm
pip install termcolor

(Time Is In-Built Module)

(OS Is In-Built Module)

'''

# Imports 

# For Browser Automation

from selenium import webdriver

from selenium.common.exceptions import NoSuchElementException

from selenium.webdriver import ActionChains

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# For Coloured CLI

from tqdm import tqdm
from termcolor import colored,cprint

# For Sending Attachments

import autoit

# For Reading Excel

import pandas 

# For Adding Buffer/Wait Time In Program

import time 

# For Copying The Message

import pyperclip 

# For Clearing CLI

import os

clear = lambda: os.system('cls')

space_count = 80 * ' '

# Program Logo

logo = '''										     _                             _  _  _ 
										    |_|   _|_ _ \    /|_  _ _|_ _ |_||_)|_)
										    | ||_| |_(_) \/\/ | |(_| |__> | ||  |  
'''

# Load Driver (Replace With Your Path) 

chromepath = r"C:\\Program Files\\Google\\Chrome Driver\\chromedriver.exe"

options = webdriver.ChromeOptions()

# Save The Session Data (THIS IS NOT SAFE) (Need To Scan QR Code Only Once For The First Time) (Replace With Your Path) 

options.add_argument("user-data-dir=C:\\Prajwal\\Python\\WhatsApp Automation\\User Data") # (For Security You Can Comment Out This Line)

while True:

	count = 0

	# Main Menu

	cprint(logo,'green',attrs=['bold'])

	# cprint('\n{}    *** Ultimate WhatsApp Automation ***'.format(space_count),'red',attrs=['bold'])

	cprint('\n{}*** Programmed By Prajwal Vijaykumar Mali ***'.format(space_count),'green',attrs=['bold'])

	cprint('\n\t1. Press 1 If You Want To Send A Message','cyan', attrs=['bold'])	

	cprint('\n\t2. Press 2 If You Want To Send An Attachment','yellow', attrs=['bold'])

	cprint('\n\t3. Press 3 If You Want To Send Both A Message And An Attachment','magenta', attrs=['bold'])

	cprint('\n\t4. Press 4 If You Want To Exit','red', attrs=['bold'])

	cprint('\n\t~Ultimate WhatsApp Automation : ','green', attrs=['bold'])

	choice = int(input('\n\t'))   

	clear() 

	# For Sending Message

	if choice == 1:

		# To Open Automated Browser Window

		driver = webdriver.Chrome(executable_path = chromepath, options = options)

		# Wait For 1 Second

		time.sleep(1)

		# To Maximize Automated Browser Window

		driver.maximize_window()
		
		# Open WhatsApp URL In Chrome Browser

		driver.get("https://web.whatsapp.com/")

		wait = WebDriverWait(driver, 60)

		# Read Data From Excel (Replace With Your Path) 

		excel_data = pandas.read_excel("C:\\Prajwal\\Python\\WhatsApp Automation\\demo.xlsx")

		# Iterate Excel Rows From Start To End

		for column in excel_data['Name'].tolist():

			# Assign Customized Message

			''' 

			1st Line Of Code Is To Send One Message Dynamically To List Of Contacts 
			2nd Line Of Code Is To Send Separate Dynamic Message To Respective Contact In The Name Column
    
			*** IMPORTANT *** - Comment Out Anyone Of The Line

			'''

			message = excel_data['Message'][0]

			# message = excel_data['Message'][count]
    
			# Locate Search Box Through x_path

			search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'

			person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

			# Clear Search Box If Any Contact Number Is Written In It

			person_title.clear()

			# Wait For 1 Second

			time.sleep(1)

			# Send Contact Number In Search Box

			person_title.send_keys(Keys.HOME)

			person_title.send_keys(str(excel_data['Contact'][count]))

			count += 1

			# Wait For 1 Second To Search Contact Number

			time.sleep(1)
		
			try:

				'''
		
				Load Error Message If Contact Number Is Unavailable/Wrong 
				(Currently Only Saved Contact Numbers Are Supported BUT, 
				You Can Use This Program To AUTOMATE Message For Those Unsaved Contacts 
				To Whom You Have Texted Atleast One Message Like HI Or Anything Else)

				'''

				element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')

			except NoSuchElementException:

				'''

				Format The Message From Excel 

				(Make The Message Dynamic Or Specific To The Receiver By Adding His/Her/Others Name. 
				You Can Also Make It Specific In Your Own Way By Adding Other Data Instead Of Name 
				Such As Company Name, College Name, Location, Etc. MORE DATA MORE SPECIFIC)

				'''

				message = message.replace('{name}', column)

				# Copy Message To Clipboard

				pyperclip.copy(message)

				person_title.send_keys(Keys.ENTER)

				actions = ActionChains(driver)

				# Paste Message (CTRL + V)

				actions.key_down(Keys.CONTROL).send_keys('v') 

				actions.perform()

				# Wait For 1 Second

				time.sleep(1)

				# Locate Send Button Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button').click()

				# Wait For 1 Second

				time.sleep(1)

		# To Clear CLI

		clear()

		# Close Chrome Driver 

		driver.quit()

		cprint('\n\tDo You Want To Run The Program Again ? (y/n) : ','red',attrs=['bold'])

		again_choice  = input('\n\t')

		if (again_choice.lower() == 'y'):

			clear()

			continue

		else:

			exit()

	# For Sending Attachment

	elif choice == 2:

		cprint(logo,'green',attrs=['bold'])

		# cprint('\n{}    *** Ultimate WhatsApp Automation ***'.format(space_count),'red',attrs=['bold'])

		cprint('\n{}*** Programmed By Prajwal Vijaykumar Mali ***'.format(space_count),'green',attrs=['bold'])

		cprint('\n\tEnter the Attachment File Path You Want To Send : ','cyan', attrs=['bold'])

		cprint('\n\t~Ultimate WhatsApp Automation : ','green', attrs=['bold'])
			
		# Provide Attachment Path

		attachment_path = input('\n\t')

		clear()

		# To Open Automated Browser Window

		driver = webdriver.Chrome(executable_path = chromepath, options = options)

		# Wait For 1 Second

		time.sleep(1)

		# To Maximize Automated Browser Window

		driver.maximize_window()

		# Open WhatsApp URL In Chrome Browser

		driver.get("https://web.whatsapp.com/")

		wait = WebDriverWait(driver, 60)

		# Read Data From Excel (Replace With Your Path) 

		excel_data = pandas.read_excel("C:\\Prajwal\\Python\\WhatsApp Automation\\demo.xlsx")

		# Iterate Excel Rows From Start To End

		for column in excel_data['Name'].tolist():
    
			# Locate Search Box Through x_path

			search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'

			person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

			# Clear Search Box If Any Contact Number Is Written In It

			person_title.clear()

			# Wait For 1 Second

			time.sleep(1)

			# Send Contact Number In Search Box

			person_title.send_keys(Keys.HOME)

			person_title.send_keys(str(excel_data['Contact'][count]))

			count += 1

			# Wait For 1 Second To Search Contact Number

			time.sleep(1)
		
			try:
		
				# Load Error Message If Contact Number Is Unavailable/Wrong 

				element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')

			except NoSuchElementException:

				person_title.send_keys(Keys.ENTER)

				actions = ActionChains(driver)

				# Locate Clip Button Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/div/span').click()

				# Wait For 1 Second

				time.sleep(1)

				# Locate Button To Send Attachment Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[1]/button').click()

				# Wait For 1 Second

				time.sleep(1)

				# To Send Attachment

				autoit.control_focus("Open", "Edit1")

				autoit.control_set_text("Open", "Edit1", attachment_path)

				autoit.control_click("Open", "Button1")

				# Wait For 1 Second

				time.sleep(1)

				# Locate Send Button Through x_path

				driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div').click()
				
				# Wait For 5 Seconds To Deliver The Attachment

				time.sleep(5)

		# To Clear CLI

		clear()

		# Close Chrome Driver 

		driver.quit() 

		cprint('\n\tDo You Want To Run The Program Again ? (y/n) : ','red',attrs=['bold'])

		again_choice  = input('\n\t')

		if (again_choice.lower() == 'y'):

			clear()

			continue

		else:

			exit()

	# For Sending Both Message And Attachment

	elif choice == 3:

		cprint(logo,'green',attrs=['bold'])

		# cprint('\n{}    *** Ultimate WhatsApp Automation ***'.format(space_count),'red',attrs=['bold'])

		cprint('\n{}*** Programmed By Prajwal Vijaykumar Mali ***'.format(space_count),'green',attrs=['bold'])

		cprint('\n\tEnter the Attachment File Path You Want To Send : ','cyan', attrs=['bold'])

		cprint('\n\t~Ultimate WhatsApp Automation : ','green', attrs=['bold'])

		# Provide Attachment Path

		attachment_path = input('\n\t')

		clear()

		# To Open Automated Browser Window

		driver = webdriver.Chrome(executable_path = chromepath, options = options)

		# Wait For 1 Second

		time.sleep(1)

		# To Maximize Automated Browser Window

		driver.maximize_window()

		# Open WhatsApp URL In Chrome Browser

		driver.get("https://web.whatsapp.com/")

		wait = WebDriverWait(driver, 60)

		# Read Data From Excel (Replace With Your Path) 

		excel_data = pandas.read_excel("C:\\Prajwal\\Python\\WhatsApp Automation\\demo.xlsx")

		# Iterate Excel Rows From Start To End

		for column in excel_data['Name'].tolist():

			# Assign Customized Message 

			''' 

			1st Line Of Code Is To Send One Message Dynamically To List Of Contacts 
			2nd Line Of Code Is To Send Separate Dynamic Message To Respective Contact In The Name Column
    
			*** IMPORTANT *** - Comment Out Anyone Of The Line

			'''

			message = excel_data['Message'][0] 

			# message = excel_data['Message'][count] 
    
			# Locate Search Box Through x_path

			search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'

			person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

			# Clear Search Box If Any Contact Number Is Written In It

			person_title.clear()

			# Wait For 1 Second

			time.sleep(1)

			# Send Contact Number In Search Box

			person_title.send_keys(Keys.HOME)

			person_title.send_keys(str(excel_data['Contact'][count]))

			count += 1

			# Wait For 1 Second To Search Contact Number

			time.sleep(1)
		
			try:

				'''
		
				Load Error Message If Contact Number Is Unavailable/Wrong 
				(Currently Only Saved Contact Numbers Are Supported BUT, 
				You Can Use This Program To AUTOMATE Message For Those Unsaved Contacts 
				To Whom You Have Texted Atleast One Message Like HI Or Anything Else)

				'''

				element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')

			except NoSuchElementException:

				'''

				Format The Message From Excel 

				(Make The Message Dynamic Or Specific To The Receiver By Adding His/Her/Others Name. 
				You Can Also Make It Specific In Your Own Way By Adding Other Data Instead Of Name 
				Such As Company Name, College Name, Location, Etc. MORE DATA MORE SPECIFIC)

				'''

				message = message.replace('{name}', column)

				# Copy Message To Clipboard

				pyperclip.copy(message) 

				person_title.send_keys(Keys.ENTER)

				actions = ActionChains(driver)

				# Paste Message (CTRL + V)

				actions.key_down(Keys.CONTROL).send_keys('v') 

				actions.perform()

				# Wait For 1 Second

				time.sleep(1)

				# Locate Send Button Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button').click()

				# Wait For 1 Second

				time.sleep(1)

				# Locate Clip Button Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/div/span').click()

				# Wait For 1 Second

				time.sleep(1)

				# Locate Button To Send Attachment Through x_path

				driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[1]/button').click()

				# Wait For 1 Second
				
				time.sleep(1)

				autoit.control_focus("Open", "Edit1")

				autoit.control_set_text("Open", "Edit1", attachment_path)

				autoit.control_click("Open", "Button1")

				# Wait For 1 Second
				
				time.sleep(1)

				# Locate Send Button Through x_path

				driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div').click()

				# Wait For 5 Seconds To Deliver The Attachment
				
				time.sleep(5)

		# To Clear CLI

		clear()

		# Close Chrome Driver 

		driver.quit() 

		cprint('\n\tDo You Want To Run The Program Again ? (y/n) : ','red',attrs=['bold'])

		again_choice  = input('\n\t')

		if (again_choice.lower() == 'y'):

			clear()

			continue

		else:

			exit()

	# To Close The Program

	elif choice == 4:

		exit()

	# When User Input Has An Error			

	else:
		
		cprint('\n\tPlease select a valid option.....!','red',attrs=['bold'])

		continue

# END
