import json
import openpyxl
from openpyxl import Workbook


def main():
    ##jsonFile is the file location for the log history's json
    ##userdiscriminator is the user's unique 4 digit number
    ##excelOutput is where you want your excel file to be outputted
    ##currentDay is the first day you want the data to start
    ##UTC hour difference is how far off your time is from UTC time
    jsonFile ='C:/Users/endri/Desktop/data visualization project/Cookaloo\'s Realm - Text Channels - general [112329740093771776].json'
    userdiscriminator = "4512"
    excelOutput = 'C:/Users/endri/Desktop/Git/Discord Data Visualization/Data.xlsx'
    currentDay="2015-06-11"
    utcHourDifference = 5

    with open(jsonFile) as a:

     j = json.load(a)
    messages = j["messages"]

    myData ={}
    hour = 0
    numMes = 0
    numEdit = 0
    numAttach = 0

    for x in messages:
        if(x["author"]["discriminator"]==userdiscriminator):

            hour = int(str(x["timestamp"])[11:13])

            #checks to see if its the same day, if it is not the same day it records the data to a list
            if((currentDay != str(x["timestamp"])[0:10] and hour >=utcHourDifference) or (int(str((currentDay)[-2:len(currentDay)]))+2<=int(str(x["timestamp"])[8:10]))):
                myData[currentDay] = [numMes,numEdit,numAttach]
                currentDay = str(x["timestamp"])[0:10]
                numMes = 1
                numEdit = 0
                numAttach = 0

                if str(x["timestampEdited"]) != "None":
                    numEdit = numEdit+1
                if str(x["attachments"]) != "[]":
                    numAttach = numAttach + 1
            else:
                numMes = numMes+1
                if str(x["timestampEdited"]) != "None":
                    numEdit = numEdit+1
                if str(x["attachments"]) != "[]":
                    numAttach = numAttach + 1
    for r in myData:
        print (str(r)+str(myData[r]))

    #creates an excel file of the data
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "WeekDay"
    sheet["B1"] = "Date"
    sheet["C1"] = "# Messages"
    sheet["D1"] = "# of Edits"
    sheet["E1"] = "# of Attachments"
    s = 2
    for e in myData:
        sheet["B"+str(s)] = "=DATE("+str(e)[0:4]+","+str(e)[5:7]+","+str(e)[8:10]+")"
        ##print(str(e))
        ##print()
        sheet["A"+str(s)] = "=TEXT(B"+str(s)+",\"dddd\")"
        sheet["C"+str(s)] = myData[e][0]
        sheet["D"+str(s)] = myData[e][1]
        sheet["E"+str(s)] = myData[e][2]
        s=s+1

    file = excelOutput
    workbook.save(file)

if __name__ == "__main__":
    main()