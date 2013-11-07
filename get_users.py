import ast
from xlrd import open_workbook
from max_surf_finder_lane import find_surf_at_spot
import datetime
import time
import imaplib
import email
from bs4 import BeautifulSoup
from email.parser import HeaderParser
import smtplib
import os
from xlwt import *
from xlrd import *
from xlutils import *
from xlutils.copy import copy
import re
import requests
import sys



def get_users():
    book = open_workbook("surf_users.xls")
    sheet = book.sheet_by_index(0)
    for i in range(0, sheet.nrows):
        contact = str(sheet.cell(i, 0).value)
        spot_list = []
        for j in range(1, sheet.ncols):            
            if sheet.cell(i, j).value != "":
                spot_list.append(ast.literal_eval(sheet.cell(i, j).value))
        find_surf_at_spot(contact, spot_list)

def one_user(number, distance=10, start_distance=0):
    book = open_workbook("surf_users.xls")
    sheet = book.sheet_by_index(0)
    i = 0
    spot_list = []
    for el in range(0, sheet.nrows):
        if sheet.cell(el, 0).value == number:
            break
        else:
            i += 1
    if i == sheet.nrows:
        msg = "Sorry, but it seems that you do not have any surf spots stored."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg)
        return
        
    for j in range(1, sheet.ncols):
        if sheet.cell(i, j).value != "":
            spot_list.append(ast.literal_eval(sheet.cell(i, j).value))
    if len(spot_list) > 0:
        find_surf_at_spot(number, spot_list, distance, start_distance)

def extract_body(payload):
    if isinstance(payload, str):
	return payload
    else:
	return '\n'.join([extract_body(part.get_payload()) for part in payload])

def get_messages(run_time):
    info = {}

    try:
        conn=imaplib.IMAP4_SSL('imap.gmail.com')
        conn.login("surfnotification","password")
        conn.select()
        typ, data = conn.search(None, 'UNSEEN')
        while True:
            if len(data[0]) != 0:
                break
            if datetime.datetime.now().hour == run_time.hour and datetime.datetime.now().minute + 10 > run_time.minute:
                break
            if run_time.minute < 10 and (datetime.datetime.now().hour == run_time.hour-1 and datetime.datetime.now().minute > 50):
                break
            conn.select()
            typ, data = conn.search(None, 'UNSEEN')
            time.sleep(.5)
    except Exception as e:
        print("error here, line 82")
        msg = "Ryan, SurfNotification has crashed. Please check you error messages."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", "<owner's phone number>@mms.att.net", msg)
        raise e
    try:
	for num in data[0].split():
            typ, msg_data = conn.fetch(num, '(RFC822)')
	    email_message = email.message_from_string(msg_data[0][1])
	    attachment = False
	    attachment_part = ""
	    for part in email_message.walk():
                if part.get_filename() != None:
                    attachment = True
                    attachment_part = part.get_payload()
                    print(attachment_part)

	    dat = conn.fetch(num,'(BODY[HEADER])')
	    header_data = dat[1][0][1]
            parser = HeaderParser()
	    person = str(parser.parsestr(header_data)["Return-Path"])[1:-1]
	    if attachment:
                message1 = make_message(attachment_part)
                info[person] = message1.strip()
                surf_parser(person, info[person])
                typ, response = conn.store(num, '+FLAGS', r'(\Seen)')
                return    
            contype = str(parser.parsestr(header_data)["Content-Type"])
	    textcontype = 'text/plain; charset="us-ascii"'
	    for response_part in msg_data:
                if isinstance(response_part, tuple):
	            msg = email.message_from_string(response_part[1])
	        payload=msg.get_payload()
	        body=extract_body(payload)
	                
	        body1 = BeautifulSoup(body)
	        if not re.compile('[0-9]{10}').match(person[0:10]):
                    msg1 = "\nSorry, but you cannot use a conventional email account. "+\
                           "Please try again with your cell phone's MMS service."
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    server.login("username@gmail.com", "password")
                    server.sendmail("SurfNotification", person, msg1)
                    return
	        message = ""
	        if attachment:
                    message1 = make_message(attachment_part)
                    break
                elif contype.strip() == textcontype:
                    message1 = make_message(body)
	        else:
                    for el in body1.find_all("td"):
                        message += str(el.get_text().strip())
                    message1 = make_message(message)
                    
	        info[person] = message1.strip()
	        print(info[person])
	    surf_parser(person, info[person])
            typ, response = conn.store(num, '+FLAGS', r'(\Seen)')
    finally:
	try:
            conn.close()
	except:
            pass
	conn.logout()

def make_message(word):
    word2 = ""
    for el in word:
        if el == "\n":
            break
        word2 += el
    word_list = word2.split()    
    final_word = ""
    for el in range(0, len(word_list)):
        final_word += word_list[el]
        if el != len(word_list) - 1:
            final_word += " "
    return final_word

def surf_parser(number, words):
    if words.lower() == "get forecast":
        one_user(number)
        return
    if words.lower() == "recommended spots":
        recommended_spots(number)
        return
    if re.compile('[0-9]{1,2} *day(s?)').match(words.lower()):
        days = int(words.split()[0])
        if days > 10:
            days = 10
        one_user(number, days)
        return
    if re.compile('[0-9]{1,2}\\-[0-9]{1,2} *day(s?)').match(words.lower()):
        lst = words.split()
        lst_nums = lst[0].split("-")
        day1 = int(lst_nums[0])
        day2 = int(lst_nums[1])
        if day2 > 10:
            day2 = 10
        one_user(number, day2, day1)
        return
    if words.lower() == "remove":
        remove(number)
        return
    if words.lower() == "commands":
        commands(number)
        return
    if words.lower().strip() == "info":
        msg = "\nWelcome to SurfNotification! This service will send you one text per day, "+\
              "notifying you of the best surf times for all your favorite spots. To get started, all you need to "+\
              "do is simply text your favorite surf spot to this address. For "+\
              "a full list of spots, please visit http://magicseaweed.com/site-map.php. "+\
              "To get a full list of commands, including spots recommended for you, text COMMANDS. Enjoy!"
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg)
        return
    original = open_workbook("surf_users.xls")
    sheet = original.sheet_by_index(0)
    spot_name = words
    req = requests.get("http://magicseaweed.com/site-map.php")
    data = req.text
    soup = BeautifulSoup(data)
    my_word_list = words.split()
    words = ""
    for el in range(0, len(my_word_list)):
        words += my_word_list[el]
        if el != len(my_word_list) - 1:
            words += " "
    words = words.strip()
    my_reg_ex = ""
    for el in words:
        my_reg_ex += "[" + el.lower() + el.upper() + "]"
    my_tag = soup.find("a", text=re.compile(my_reg_ex))
    if my_tag == None:
        msg = "\nSorry, but we couldn't find " + words + \
              ". Be sure to capitalize the first letter of major " + \
              "words. Try visiting magicseaweed.com/site-map.php for a " + \
              "full list of available surf spots. Be sure to enter them " + \
              "exactly as they appear."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg)
        return
    my_ref = my_tag.get("href")
    print(my_ref)
    r = requests.get("http://www.magicseaweed.com" + my_ref)
    data = r.text
    souper = BeautifulSoup(data)
    try:
        spot_name = souper.find("span", text=re.compile(my_reg_ex)).get_text()
    except AttributeError as e:
        print(words)
        print(my_reg_ex)
        raise e
    if len(souper.find_all("td", class_="msw-fc-s")) == 0:
        msg1 = "Unfortunately, the " + words + " report does not " + \
               "actually contain surf data. It will not be added to your list. " +\
               "Feel free to make another selection."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg1)
        return
    words = '("'+spot_name+'", "'+my_ref+'")'

    for i in range(0, sheet.nrows):
        if sheet.cell(i, 0).value == number:
            part_of_data = False
            j = 1
            for el in range(1, sheet.ncols):
                if sheet.cell(i, j).value == words:
                    part_of_data = True
                if sheet.cell(i, j).value != "":
                    j += 1
                else:
                    break
            if not part_of_data:
                msg2 = "We've received your text. " + spot_name + \
                       " will be added to your list of spots."
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login("username@gmail.com", "password")
                server.sendmail("SurfNotification", number, msg2)
                updated = copy(original)
                updated.get_sheet(0).write(i, j, words)
                updated.save("temp_surfers.xls")
                os.remove("surf_users.xls")
                updated.save("surf_users.xls")
                return
            else:
                msg3 = spot_name + " is already in your list of spots."
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login("username@gmail.com", "password")
                server.sendmail("SurfNotification", number, msg3)
                return
    updated = copy(original)
    if sheet.nrows != 0:
        msg2 = "We've received your text. " + spot_name + \
               " will be added to your list of spots."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg2)
        updated.get_sheet(0).write(sheet.nrows, 0, number)
        updated.get_sheet(0).write(sheet.nrows, 1, words)
    else:
        updated.get_sheet(0).write(0, 0, number)
        updated.get_sheet(0).write(0, 1, words)
        msg2 = "We've received your text. " + spot_name + \
                " will be added to your list of spots."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg2)
    updated.save("temp_surfers.xls")
    os.remove("surf_users.xls")
    updated.save("surf_users.xls")
    return

def commands(number):
    msg = "\nHere is a list of commands:\nREMOVE will remove you from the mailing list.\n"+\
          "GET FORECAST will send you the report for your spots for the next 10 days.\n"+\
          "X DAY(S) will send you the report for your spots for the next X days.\n"+\
          "RECOMMENDED SPOTS will send you a list of spots recommended for you.\n"
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("username@gmail.com", "password")
    server.sendmail("SurfNotification", number, msg)

area_map = {"925":"Central California", "510":"Central California", "650":"Central California", "415":["Central California", "Northern California"],
            "831":"Central California", "669":"Central California", "408":"Central California", "209":"Central California",
            "559":"Central California", "707":"Northern California", "530":"Northern California", "916":"Northern California"}

def recommended_spots(number):
    area_code = number[0:3]
    try:
        area = area_map[area_code]
    except KeyError:
        msg = "Sorry, but we do not have your area."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", number, msg)
        return
    r = requests.get("http://magicseaweed.com/site-map.php")
    data = r.text
    soup = BeautifulSoup(data)
    thing = soup.find("h1", class_="header", text=re.compile(area + " Surf Reports"))
    table = thing.find_next_sibling("table")
    spot_list = []
    for el in table.find_all('a'):
        spot_list += [(el.get_text()),]
    msg = "\nHere is a list of spots that may pertain to your area:\n"
    for el in range(0, len(spot_list)):
        msg += spot_list[el]
        if el != len(spot_list) - 1:
            msg += ", "
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("username@gmail.com", "password")
    server.sendmail("SurfNotification", number, msg)
                          
    

def remove(number):
    original = open_workbook("surf_users.xls")
    sheet = original.sheet_by_index(0)
    updated = copy(original)
    for i in range(0, sheet.nrows):
        if sheet.cell(i, 0).value == number:
            for el in range(i+1, sheet.nrows):
                for j in range(0, sheet.ncols):
                    updated.get_sheet(0).write(el-1, j, sheet.cell(el, j).value)
            for el in range(0, sheet.ncols):
                updated.get_sheet(0).write(sheet.nrows-1, el, "")
    updated.save("temp_surfers.xls")
    os.remove("surf_users.xls")
    updated.save("surf_users.xls")
    msg4 = "We've received your text and we're sad to see you go. " + \
           "You will not receive any more texts."
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("username@gmail.com", "password")
    server.sendmail("SurfNotification", number, msg4)
    return

run_time = datetime.time(17, 0)
while True:
    try:
        right_now = datetime.datetime.now()
        time.sleep(.1)
        get_messages(run_time)
        if (right_now.hour == run_time.hour and right_now.minute == run_time.minute):
            get_messages(run_time)
            get_users()
            time.sleep(70)
    except KeyboardInterrupt as e:
        raise e
    except BaseException as e:
        time.sleep(10)
        msg = "Ryan, SurfNotification has crashed. Please check you error messages."
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("username@gmail.com", "password")
        server.sendmail("SurfNotification", "<owner's phone number>@mms.att.net", msg)
        tb = sys.exc_info()[2]
        print("Line number: " + str(tb.tb_lineno))
        print("Time: " + str(datetime.datetime.now()))
        print(e)
        raise e    
