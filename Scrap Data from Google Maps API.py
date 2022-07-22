import googlemaps

def GDistMat(reqtime, origins, destinations, timestamp):
    #Google Maps Travel Time Output Tool
    #No User Inputs Required

    import os
    import datetime
    import calendar
    import numpy
    import pprint
    import time
    maps_key = "AIzaSyCec1pahrUuYKkoAcCucbUAwBPST6Tf7Z8"
    gmaps = googlemaps.Client(key=maps_key)

    deptime = reqtime

    #Pull information from the Google distance API. This currently pulls based on their "best guess algorithm"
    try:
        #Best Guess
        BestGuess = gmaps.distance_matrix(
            (origins),
            (destinations),
            departure_time = deptime,
            mode = 'driving',
            traffic_model = 'best_guess',
            )

        # BestGuess is a json of all the data pulled by the API
        #   Indexes json from google api to rows.
        data = BestGuess['rows']
        #Iterates through each "elements" dictionary in "data"
        travel_list = {}
        travel_list['rtime'] = time.strftime("%H:%M:%S")
        for i in data:
            #Populates a dictionary called "Travel_list "
            travel_list['node1'] = BestGuess['origin_addresses'][0]
            travel_list['node2'] = BestGuess['destination_addresses'][0]
            travel_list['time'] = i['elements'][0]['duration_in_traffic']['value']
            travel_list['distance'] = i['elements'][0]['distance']['value']
        #print("yay") #debug
    except Exception as e:
        #print("error")
        print(e)

        travel_list = {}
        travel_list['rtime'] = time.strftime("%H:%M:%S")
        #for i in data:
        #Populates a dictionary called "Travel_list "
        travel_list['node1'] = ""
        travel_list['node2'] = ""
        travel_list['time'] = 999999
        travel_list['distance'] = 999999


    #Returns the Travel_list dictionary to be used in function.
    return travel_list


# In[ ]:

# In[ ]:


def csv_read(filename):
    import sys
    import time
    import csv
    import datetime
    #import schedule
    import pandas as pd

    #--------------------------------------------------------------------------------

    #Reads CSV with list of origins and destinations
    try:
        OD_list = pd.read_csv(filename, encoding='iso-8859-1')

        origin_list = []
        destination_list = []
        origin_list = OD_list['origin'].tolist()
        destination_list = OD_list['destination'].tolist()
    except ():
        print('Appropriate CSV not chosen')
#Returns lists of origins, destinations, and descriptions.
    return origin_list, destination_list

#-------------------------------------------------
def gettimenow():

    timer = datetime.datetime.today()
    dateref = timer.date()

    day = dateref.strftime("%A")
    date = dateref.strftime("%d")
    month = dateref.strftime("%B")
    time = datetime.datetime.utcnow()

    return time


def apicall(origin_list, destination_list):
    #Calls google API caller, requests data for the list of OD pairs and prints that data into the table created in SQL.
    #global rowCounter
    #import psycopg2
    import time
    import csv
    import pprint
    import datetime
    from datetime import date, timedelta
    import numpy as np
    #import schedule

    #Define when you want the script to stop by time or by request limit. Time is in 24:00 format.
    #houston("initialize")

    for i in range(0, len(origin_list)):
        #Defines which OD pair to pull from the long list
        origins = origin_list[i]
        destinations = destination_list[i]

        reqtime= "now"
        timestamp = time.strftime("%H:%M:%S")
        travel_info = GDistMat(reqtime, origins, destinations, timestamp)
        #Increase requests number by 1 for each request made. Stops the code if the requests gets above the limit.
        #requests += 1



    return travel_info

from tkinter import Tk
from tkinter.filedialog import askopenfilename
#import sys
#filename = sys.argv[1]
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)


origin_list, destination_list = csv_read(filename)

#global rowCounter
#import schedule
import datetime
import time
import xlsxwriter
outWorkbook = xlsxwriter.Workbook("output.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1", "Run_time")
outSheet.write("B1", "time")
delaytime = 10
repeat = 3
for i in range(0, repeat):
    starttime = gettimenow()
    travel_list = apicall(origin_list, destination_list)
    travel_list['day'] = starttime.day
    travel_list['date'] = str(starttime.date)
    travel_list['month'] = starttime.month
    #year = datetime.datetime.today()
    #year = year.year
    d1_list=list(travel_list.values())
    outSheet.write(i+1, 0, d1_list[0])
    outSheet.write(i+1, 1, d1_list[3])
    print(travel_list)
    endtime = gettimenow()
    calltime = endtime - starttime
    time.sleep(delaytime - calltime.total_seconds())
outWorkbook.close()
