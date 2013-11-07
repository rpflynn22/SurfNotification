from bs4 import BeautifulSoup
import requests
import datetime
import time
import smtplib

times = ["12am","3am","6am","9am","12pm","3pm","6pm","9pm"]

def find_surf_at_spot(number, info, distance=10, start_distance=0):
    if start_distance != 0:
        between = True
    else:
        between = False
    days = []
    best_days = []
    max_surf_overall = 0
    max_surf_spot = ""
    max_surf_day_overall = ""
    for place in info:
        table = []
        r = requests.get("http://www.magicseaweed.com" + place[1])
        data = r.text
        soup = BeautifulSoup(data)
        for el in soup.find_all("span", class_="msw-fc-day"):
            days.append(el.get_text())
        for link in soup.find_all("td", class_="msw-fc-s"):
            try:
                aquisition = len(link.get_text()) - 4
                my_val = link.get_text()[aquisition:aquisition+2]
                if my_val[0] == '-':
                    my_val = my_val[1]
                    table.append(int(my_val))
            except ValueError:
                pass
                
        start_index = int(round((datetime.datetime.now().hour / 3) + .5))
                
                
        for el in range(0, start_index + 1):
            table[el] = 0
        for el in range(0, start_distance * 8):
            table[el] = 0
        table = table[0:((distance * 8) + 1)]
        for el in range(0, len(table)):
            rem = el % 8
            if rem == 0 or rem == 1:
                table[el] = 0
                
        max_surf = max(table)
        max_surf_index = table.index(max(table))
        max_surf_day = days[int(max_surf_index / 8)]
        best_days += [max_surf_day + " at " + times[max_surf_index % 8] + " at " + place[0] + ": " + str(max_surf) + "ft.",]

    if not between:
        msg = "\nHere are the top surf times for the next " + str(distance) + " days at your locations:"
    else:
        msg = "\nHere are the top surf times from " + str(start_distance) + " days to " + str(distance) + " days at your locations:"
    for i in range(len(best_days)):
        msg += "\n" + str(best_days[i])


    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("username@gmail.com", "password")
    server.sendmail("SurfNotification", number, msg)
