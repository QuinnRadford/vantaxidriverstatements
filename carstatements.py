import sqlite3
import sys
import sqlite3.dump
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from datetime import timedelta
from datetime import timezone
from decimal import *

diskconn = sqlite3.connect('statements.db')
memconn = sqlite3.connect(':memory:')
query = "".join(line for line in diskconn.iterdump())
memconn.executescript(query)
c = memconn.cursor()
file=""
while file == "":
    print('Enter Report File Name (including file type "example.xml"): ')
    file = input()
    if file == "":
        print("No file selected")
chargesfile=""
while chargesfile == "":
    print('Enter Charges File Name (including file type "example.csv"): ')
    chargesfile = input()
    if file == "":
        print("No file selected")
print('Enter Statement Note/Reason or ENTER to skip: ')
statementnote = input()
if statementnote == "":
    statementnote = "None"
    print("Note Skipped.")

idpattern = re.compile('(\d\d*)')

optiontree = ET.parse('options.xml')
optionroot = optiontree.getroot()

#weekday names
allweekdays = {
    "0":"monday",
    "1":"tuesday",
    "2":"wednesday",
    "3":"thursday",
    "4":"friday",
    "5":"saturday",
    "6":"sunday"
    }

#get sedan lease rates
sedanrates=optionroot.find("leases").find("sedan")
sedanddaylease = sedanrates.find("day").text
nightsedanleases = {
    "0":sedanrates.find("monday").text,
    "1":sedanrates.find("tuesday").text,
    "2":sedanrates.find("wednesday").text,
    "3":sedanrates.find("thursday").text,
    "4":sedanrates.find("friday").text,
    "5":sedanrates.find("saturday").text,
    "6":sedanrates.find("sunday").text
    }
#get van lease rates
vanrates=optionroot.find("leases").find("van")
vanldayease = vanrates.find("day").text
nightvanleases = {
    "0":vanrates.find("monday").text,
    "1":vanrates.find("tuesday").text,
    "2":vanrates.find("wednesday").text,
    "3":vanrates.find("thursday").text,
    "4":vanrates.find("friday").text,
    "5":vanrates.find("saturday").text,
    "6":vanrates.find("sunday").text
    }
#get tcar lease rates
tcarlease = optionroot.find("leases").find("tcar").text

tree = ET.parse(file)
root = tree.getroot()

#insert xml into db
for detail in root.findall('{urn:crystal-reports:schemas:report-detail}Details'):
    car = detail[0][0][0].text
    timeonstring = detail[0][1][0].text
    driverstring = detail[0][3][0].text
    timeoffstring = detail[0][2][0].text
    if(timeonstring != '-' and timeoffstring != '-'):
        datetimeon = datetime.strptime(timeonstring, '%d %b %Y, %H:%M')
        datetimeoff = datetime.strptime(timeoffstring, '%d %b %Y, %H:%M')
        driverid = idpattern.match(driverstring).group(0)
        drivername = re.sub(r'(\d\d* \- )', ' ', driverstring)
        houron = datetimeon.hour
        houroff = datetimeoff.hour
        shiftlength = datetimeoff - datetimeon
        if houron >= 2 and houroff <= 17 and shiftlength.seconds > 7200:
            shifttype = 0
        elif houron >= 2 and houron <= 15 and shiftlength.seconds < 7200 and shiftlength.seconds > 3600:
            shifttype = 0
        elif shiftlength.seconds > 7200:
            shifttype = 1
        else:
            shifttype = 2
        
        c.execute('INSERT INTO Shifts (DriverID, Car, Start, End, Type, Date, Length, DriverName) VALUES ("' + driverid + '", "' + car + '", ' + str(int(datetimeon.timestamp())) + ', ' +str(int(datetimeoff.timestamp())) + ', ' + str(shifttype) + ', ' + str(datetimeon.year) + str(datetimeon.month).zfill(2) + str(datetimeon.day).zfill(2) + ', ' + str(shiftlength.seconds) + ', "' + str(drivername) + '");')
print("XML Okay!")

#instert charges into db
chargenumber = 0
with open(chargesfile) as charges:
    #read line by line
    for cnt, line in enumerate(charges):
        chargevalues = line.split(",")
        if cnt == 0:
            tablecols = ','.join(str('"' + x + '"') for x in chargevalues)
            chargeheader = chargevalues
            colnames = []
            for header in chargeheader:
                colnames.append('"' + header + '"' + ' numeric')
            c.execute('CREATE TABLE Extras (' + ','.join(str(x) for x in colnames) + ');')
        elif chargevalues[0].isdigit():
            c.execute('INSERT INTO Extras (' + tablecols + ') values(' + line + ');')
            c.execute('INSERT INTO Charges (DriverID, Charge, Credit) VALUES (' + chargevalues[0] + ', ' + chargevalues[1] + ', ' + chargevalues[2] + ')')
            chargenumber += 1
if chargenumber > 0:
    print(str(chargenumber) + " Charges Okay!")
else:
    print("No Charges!")
        

print("Writing Driver File")

driverfile = 'driver_' + file.replace('.', '_').replace('/', '_') + '.csv'
driverdata = open(driverfile, 'w')

csvheader = 'DriverID, Car, Type, Date, Length, Value, DriverName \n'

#lease values

cartypes = {
    "1":0,
    "2":0,
    "3":0,
    "4":0,
    "5":0,
    "6":0,
    "7":0,
    "8":0,
    "9":0,
    "10":0,
    "11":0,
    "12":0,
    "13":0,
    "15":0,
    "16":0,
    "17":0,
    "18":0,
    "20":0,
    "22":0,
    "23":0,
    "25":0,
    "26":0,
    "29":0,
    "30":1,
    "31":1,
    "32":1,
    "33":1,
    "34":1,
    "35":1,
    "36":1,
    "37":1,
    "38":1,
    "39":1,
    "40":1,
    "41":1,
    "42":1,
    "43":1,
    "44":1,
    "45":1,
    "46":1,
    "47":1,
    "48":1,
    "49":1,
    "50":1,
    "51":1,
    "53":1,
    "54":1,
    "55":1,
    "56":1,
    "57":1,
    "58":1,
    "59":1,
    "60":1,
    "62":1,
    "63":0,
    "66":0,
    "67":0,
    "69":0,
    "71":1,
    "74":0,
    "76":0,
    "77":0,
    "79":1,
    "81":1,
    "88":0,
    "90":0,
    "91":0,
    "94":1,
    "98":0,
    "99":0,
    "100":0,
    "101":1,
    "102":1,
    "104":0,
    "107":0,
    "110":0,
    "111":0,
    "116":1,
    "157":0,
    "170":1,
    "200":0,
    "201":0,
    "202":0,
    "203":0,
    "204":0,
    "205":0,
    "206":0,
    "207":0,
    "208":0,
    "209":0,
    "210":0,
    "211":0,
    "212":0,
    "213":0,
    "214":0,
    "215":0,
    "216":0,
    "217":0,
    "218":0,
    "219":0,
    "220":0,
    "221":0,
    "222":0,
    "223":0,
    "224":0,
    "225":0,
    "226":1,
    "227":1,
    "228":1,
    "229":1,
    "230":1,
    "231":0,
    "700":2,
    "701":2,
    "702":2,
    "703":2,
    "704":2,
    "705":2,
    "706":2,
    "707":2,
    "708":2,
    "709":2,
    "710":2,
    "711":2,
    "712":2,
    "713":2,
    "715":2,
    "717":2,
    "719":2,
    }

#get all the dates in range
c.execute('select distinct Date from Shifts')
daterows = []
alldriverdata = ""
for row in c.fetchall():
    daterows.append(row[0])
    
#get all the drivers in range
c.execute('select distinct Car from Shifts')
carrows = []
shifnums = 0
for row in c.fetchall():
    carrows.append(row[0])

#Open Template
templatestring = ""
templatefile = open('templates/statement-template.html', 'r')
templatestring = templatefile.read()

#collapse shifts for each driver
for car in carrows:
    driverID = ""
    drivername = ""
    thistemplate = ""
    

    
    for date in daterows:
        c.execute('select Car, Start, End, Type, Length, DriverName, DriverID from Shifts where Date = ' + str(date) + ' and Car = ' + str(car) +';')
        shifttime = 0
        startdates = []
        enddates = []
        types = []
        shiftvalue = 0
        driversname = []
        output = ""
        shiftType = ""
        cartypestring = ""
        
        for row in c.fetchall() or []:
            if row[3] < 2:
                types.append(row[3])
                shifttime += row[4]
                startdates.append(row[1])
                driversname.append(row[5])
                enddates.append(row[1])
                thisdriverid.append([])
            else:
                print('shift skipped, lost ' + str(round(((row[4]/60)/60),2)) + 'h')
        
        #do write out the days data    
        if shifttime > 0:
            weekday = str(datetime.fromtimestamp(startdates[0]).weekday())
            thiscar = row[0].lstrip("V")
            #figure out the lease value
            if cartypes[thiscar] == 0:
                #car
                cartypestring = "C"
                if types[0] == 0:
                    shiftType = 'DAY'
                    shiftvalue = sedanddaylease
                elif types[0] == 1:
                    shiftType = 'NIGHT'
                    shiftvalue = nightsedanleases[weekday]
                else:
                    print("shift type error!")
            elif cartypes[thiscar] == 1:
                #van
                cartypestring = "V"
                if types[0] == 0:
                    shiftType = 'DAY'
                    shiftvalue = vanldayease
                elif types[0] == 1:
                    shiftType = 'NIGHT'
                    shiftvalue = nightvanleases[weekday]
                else:
                    print("shift type error!")
            elif cartypes[thiscar] == 2:
                #tcar
                cartypestring = "T"
                shiftType = 'TCAR'
                shiftvalue = tcarlease
            else:
                shiftType = 'ERROR'
                print("car type error!")
            output = str(driver) + ', ' + str(thiscar) + ', ' + str(types[0]) +', '+ str(date) +', '+ str(round(((shifttime/60)/60), 2)) + ', ' + str(shiftvalue) + ', ' + driversname[0] + '\n'
            rowoutput = rowoutput + '<tr><td>VT' + str(thiscar).ljust(8).replace(' ', '&nbsp;') + str(cartypestring).ljust(10).replace(' ', '&nbsp;') + str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7]).ljust(15).replace(' ', '&nbsp;') + allweekdays[weekday].ljust(12).replace(' ', '&nbsp;') + shiftType.ljust(8).replace(' ', '&nbsp;') + '</td>' + '<td class="rightalign">$' + str(shiftvalue) + '.00</td></tr>'
            driverID = str(driver)
            drivername = driversname[0]
            totalvalue = totalvalue + int(shiftvalue)
            alldriverdata += output
            shifnums +=1
    #add values to template
    thistemplate = templatestring
    thistemplate = thistemplate.replace('$STATEMENT_NOTE', str(statementnote))
    thistemplate = thistemplate.replace('$DRIVER_NAME', drivername)
    thistemplate = thistemplate.replace('$DRIVER_NUMBER', driverID)
    thistemplate = thistemplate.replace('$SHIFT_DATA', rowoutput)
    thistemplate = thistemplate.replace('$CHARGE_ROWS', chargeoutput)
    thistemplate = thistemplate.replace('$LEASE_TOTAL', str(Decimal(totalvalue).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)))
    thistemplate = thistemplate.replace('$TOTAL_BALANCE', str(Decimal(totalchargevalue - totalvalue).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    thistemplatefilename = 'statements/' + str(driver) + '_' + drivername.replace(' ', '_') + '_statement_' + file.replace('.', '_') + '.html'
    thistemplatefile = open(thistemplatefilename, 'w')
    thistemplatefile.write(thistemplate)
    thistemplatefile.close()
    

driverdata.write(csvheader + alldriverdata)
print(str(shifnums) + " shifts written")
templatefile.close()
driverdata.close()
c.close()
memconn.close()
print("all done!")
print('Press enter to close this window')
file = input()
