import sqlite3
import sys
import sqlite3.dump
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from datetime import timedelta
from datetime import timezone
from decimal import Decimal
from decimal import ROUND_HALF_DOWN
import pdfkit
from pprint import pprint

def formatCharges(chargeval):
    if chargeval < 0:
        return '($' + str(Decimal(chargeval*int(-1)).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + ')'
    else:
        return '$' + str(Decimal(chargeval).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;'

class UserInputs:
    def __init__(self):
        self.logonPath = ""
        self.chargePath = ""
        self.statementNote = ""
        self.longWeekendDate = ""
        self.longWeekendDateTime = datetime.now()
    def userInput(self, message):
        thisInput = ""
        while thisInput == "":
            print(message)
            thisInput = input()
            if thisInput == "":
                print("Invalid Response")
        return thisInput
    def userSetlogonPath(self, skip=False):
        if not skip:
            self.logonPath = self.userInput('Enter Report File Name (including file type "example.csv"): ')
        else:
            print("skipping input")
            self.logonPath = skip
    def userSetChargePath(self, skip=False):
        if not skip:
            self.chargePath = self.userInput('Enter Charges File Name (including file type "example.csv"): ')
        else:
            print("skipping input")
            self.chargePath = skip
    def userSetStatementNote(self, skip=False):
        if not skip:
            self.statementNote = self.userInput('Enter Statement Note/Reason: ')
        else:
            print("skipping input")
            self.statementNote = skip
    def userSetLongWeekendDate(self):
        while self.longWeekendDate == "":
            print('Does this month contain a long weekend? enter y/n')
            self.longWeekendDate = input()
            if self.longWeekendDate == "n":
                print("No long weekend")
            if self.longWeekendDate == "y":
                print("Enter the date of the Sunday on the long weekend dd/mm/yy")
                self.longWeekendDate = input()
                self.longWeekendDateTime = datetime.strptime(self.longWeekendDate, '%d/%m/%y')
            else:
                print("invalid response")
                self.longWeekendDateTime = "none"

class ShiftCSV:
    WEEKDAY_NAMES = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
    ]
    shiftMap = {
        0:"Day",
        1:"Night",
        2:"Tcar"
    }
    csvHeader = "account,amount,comment\n"
    def __init__(self):
        self.csvValue = ""
    def appCSV(self, lineAppend):
        self.csvValue = self.csvValue + lineAppend
    def formatDate(self, date, mode):
        if mode == 1:
            return str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7])
        else:
            return str(str(date)[4] +str(date)[5] + '-' + str(date)[6] +str(date)[7] + '-' + str(date)[2] +str(date)[3])
    def formatDriverAccount(self, driverId):
        idString = str(driverId)
        if len(idString) > 6:
            idString = "23" + idString[-6:]
        else:
            idString = "23" + idString.ljust(6,"0")
        return idString
    def formatCarAccount(self, car, shift):
        shiftString = str(int(shift + 1)) #shift up shift type by 1
        carString = str(car)
        carString = carString.rjust(3,"0") #pad car number with 0s to the right
        carString = "22" + carString.ljust(5,"0") + shiftString
        return carString
    def appendShift(self, driverId, car, date, shift, value, name):
        driverAccount = self.formatDriverAccount(driverId)
        carAccount = self.formatCarAccount(car, shift)
        weekdayId = datetime.strptime(date, '%Y%m%d').weekday()
        lineCols = {
        "driver" : [
            driverAccount,
            str(value)
        ],
        "car" : [
            carAccount,
            "-" + str(value)
        ],
        "tx" : [
            self.formatDate(date, 2),
            "lease",
            str("VT" + str(car) + " " + self.shiftMap[shift] + " " + self.WEEKDAY_NAMES[weekdayId] + " - ID " + str(driverId) + " " + name)
        ]
        }
        self.appCSV(str(','.join(lineCols["tx"]) + '\n'))
        self.appCSV(str(','.join(lineCols["driver"]) + '\n'))
        self.appCSV(str(','.join(lineCols["car"]) + '\n'))
    def writeFile(self, csvPath):
        thisDriverdata = open('csv/' + csvPath, 'w')
        thisDriverdata.write(self.csvHeader + self.csvValue)
        thisDriverdata.close()

def saveStatement(cartemplatestring, daterange, statementnote, dayNight, shiftoutput, leasetotal, car):
    carthistemplate = cartemplatestring #instantiate car template
    carthistemplate = carthistemplate.replace('$DATERANGE', daterange)
    carthistemplate = carthistemplate.replace('$STATEMENT_NOTE', str(statementnote))
    carthistemplate = carthistemplate.replace('$CAR_NAME', car + dayNight[0])
    carthistemplate = carthistemplate.replace('$SHIFT_DATA', shiftoutput)
    carthistemplate = carthistemplate.replace('$LEASE_TOTAL', str(Decimal(leasetotal).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)))
    carthistemplatefilename = 'statements/car/thiscar.html' #temporary html file for further processing into pdf
    carthistemplatefilenamepdf = 'statements/car/' + str(car) + '_NIGHT_statement_' + userInput.logonPath.replace('.', '_') + '.pdf'
    carthistemplatefile = open(carthistemplatefilename, 'w')
    carthistemplatefile.write(carthistemplate)
    carthistemplatefile.close()
    pdfkit.from_file(carthistemplatefilename, carthistemplatefilenamepdf,options = options, configuration = config) #save temp HTML file as properly named PDF

def genStatement(carrows, dayNight, carTemplatePath):
    #carTemplatePath string path to template
    #carrows array [car, shift[]]
    #dayNight str values: NIGHT, DAY
    WEEKDAY_NAMES = [
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday"
    ]

    STATEMENT_HEADER = {
    "DRIVER" : 36,
    "DATE" : 15,
    "SHIFT" : 12
    } #data : length

    #Table structure templates
    HEADER_STRUCT = '<tr><td>$CONTENT</td><td></td></tr>'
    TABLE_STRUCT = '<tr><td>$CONTENT</td><td style="text-align:right;">$VALUE</td></tr>'

    progress = 0 #Progress counter
    totalstatements = len(carrows) #total of

    cartemplatefile = open(carTemplatePath, 'r')
    cartemplatestring = cartemplatefile.read() #read the car template

    headerItems = ""
    for head, length in STATEMENT_HEADER.items():#build the header line content
        headerItems += head.ljust(length).replace(' ', '&nbsp;')
    
    headerLine = HEADER_STRUCT.replace('$CONTENT', headerItems)

    for car, shifts in carrows.items():
        leasetotal = 0 #initialize total lease
        shiftoutput = headerLine #initialize output with table header line
        for shift in shifts:
            if shift['shift'] == dayNight or shift['shift'] == 'TCAR':
                leasetotal += int(shift['value'])
                weekdayId = datetime.strptime(shift['date'], '%Y/%m/%d').weekday()
                thisWeekdayName = WEEKDAY_NAMES[weekdayId]
                shiftCols = {
                    shift['driver'] : 8,
                    shift['name'] : 28,
                    shift['date'] : 15,
                    thisWeekdayName : 12
                }
                colItems = ""
                for col, length in shiftCols.items(): #concat data columns
                    colItems += col.ljust(length).replace(' ', '&nbsp;')
                thisShiftRow = TABLE_STRUCT
                thisShiftRow = thisShiftRow.replace('$CONTENT',colItems)
                thisShiftRow = thisShiftRow.replace('$VALUE',shift['value'] + '.00')
                shiftoutput += thisShiftRow #append row to output table

        saveStatement(cartemplatestring, daterange, userInput.statementNote, dayNight, shiftoutput, leasetotal, car)
        print("Written car statement " + dayNight + " " + str(progress) + " of " + str(totalstatements))
        progress += 1

class leaseOptions:
    def __init__(self, optionPath):
        self.optiontree = ET.parse(optionPath)
        self.optionroot = self.optiontree.getroot() #open options xml file for lease rates
        self.sedanrates=self.optionroot.find("leases").find("sedan")
        self.vanrates=self.optionroot.find("leases").find("van")
        self.sedanLeases = {
            "0":self.sedanrates.find("monday").text,
            "1":self.sedanrates.find("tuesday").text,
            "2":self.sedanrates.find("wednesday").text,
            "3":self.sedanrates.find("thursday").text,
            "4":self.sedanrates.find("friday").text,
            "5":self.sedanrates.find("saturday").text,
            "6":self.sedanrates.find("sunday").text,
            "day":self.sedanrates.find("day").text
        }
        self.vanLeases = {
            "0":self.vanrates.find("monday").text,
            "1":self.vanrates.find("tuesday").text,
            "2":self.vanrates.find("wednesday").text,
            "3":self.vanrates.find("thursday").text,
            "4":self.vanrates.find("friday").text,
            "5":self.vanrates.find("saturday").text,
            "6":self.vanrates.find("sunday").text,
            "day":self.vanrates.find("day").text
        }
        #get tcar lease rates
        self.tcarlease = self.optionroot.find("leases").find("tcar").text
    def getSedanRate(self, day):
        return self.sedanLeases[day]
    def getVanRate(self, day):
        return self.vanLeases[day]
    def getTcarRate(self):
        return self.tcarlease

def setExemptDrivers(con, ownerConfigPath):
    #get car owners
    con.execute('CREATE TABLE Exemptions (driverID, car, shift);')
    with open(ownerConfigPath) as ownerids:
        #read line by line
        for cnt, line in enumerate(ownerids):
            owneridvals = line.split(",")
            if cnt == 0:
                print('Reading Owner IDs')
            else:
                con.execute('INSERT INTO Exemptions (driverID, car, shift) VALUES (' + owneridvals[1] + ', ' + owneridvals[0] + ', ' + owneridvals[3] + ')')

#Setup initial data
idpattern = re.compile(r'(\d\d*)') #compile regex for parsing driver ID

#PDFkit global Constants
path_wkthmltopdf = r'wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
options = {
    'quiet': ''
    }
#SQL Connection
diskconn = sqlite3.connect('config/statements.db')
memconn = sqlite3.connect(':memory:')
query = "".join(line for line in diskconn.iterdump())
memconn.executescript(query)
c = memconn.cursor()
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
setExemptDrivers(c,'config/owner_id.csv')
leaseValues = leaseOptions('config/options.xml')
userInput = UserInputs()
userInput.userSetlogonPath("48-charge.csv")
userInput.userSetChargePath("charges.csv")
userInput.userSetStatementNote("none")
userInput.userSetLongWeekendDate()
shiftReport = ShiftCSV()

#Service Hours Log
serviceHours = 'Logon,Logoff,Driver,ID,Length in hours\r'
#End setup phase

#insert csv into db
with open(userInput.logonPath) as shiftrows:
    for cnt, line in enumerate(shiftrows):
        shifttype = 2
        houron = 0
        houroff = 0
        thisshiftrow = line.split(r'","')
        car = thisshiftrow[6]
        timeonstring = thisshiftrow[8]
        driverstring = thisshiftrow[7]
        timeoffstring = thisshiftrow[9]
        drivername = re.sub(r'(\d\d* \- )', ' ', driverstring)
        if(timeonstring != '-') and (timeonstring != '0'):
            datetimeon = datetime.strptime(timeonstring, '%d %b %Y, %H:%M')
            if(timeoffstring == '-') or (timeoffstring == '0'):
                datetimeoff = datetimeon + timedelta(hours=6)
            else:
                datetimeoff = datetime.strptime(timeoffstring, '%d %b %Y, %H:%M')
            driverid = idpattern.match(driverstring).group(0)
            houron = datetimeon.hour
            houroff = datetimeoff.hour
            shiftlength = datetimeoff - datetimeon

            #Determine whether the shift is day or night
            if houron > houroff and shiftlength.seconds < 46800:
                shifttype = 1 #night
            elif houron <= 3 and houroff <= 5 and shiftlength.seconds > 7200:
                #shift to yesterday
                datetimeon = datetimeon - timedelta(days=1)
                datetimeoff = datetimeoff - timedelta(days=1)
                shifttype = 1 #night
            elif houron >= 0 and houron <= 9 and shiftlength.seconds > 46800:
                serviceHours += str(datetimeon) + ',' + str(datetimeoff) + ','  + drivername + ',' + driverid + ',' + str(shiftlength.seconds/3600) + '\r'
                print('Illegal Hours Of Service on ' + str(datetimeon) + ' for ' + drivername + ' after ' + str(shiftlength.seconds/3600) + 'hours')
                shifttype = 0 #day
            elif shiftlength.seconds > 46800:
                serviceHours += str(datetimeon) + ',' + str(datetimeoff) + ',' + drivername + ',' + driverid + ',' + str(shiftlength.seconds/3600) + '\r'
                print('Illegal Hours Of Service on ' + str(datetimeon) + ' for ' + drivername + ' after ' + str(shiftlength.seconds/3600) + 'hours')
                shifttype = 1 #night
            elif houron <= 6 and houroff <= 6:
                shifttype = 3 #toss
            elif houron <= 6 and houroff >= 6 and shiftlength.seconds > 7200:
                shifttype = 0 #day
            elif houron >= 3 and houron <= 14 and houroff <= 17 and shiftlength.seconds > 7200:
                shifttype = 0 #day
            elif shiftlength.seconds > 7200:
                shifttype = 1 #night
            else:
                shifttype = 2 #toss
            
            c.execute('INSERT INTO Shifts (DriverID, Car, Start, End, Type, Date, Length, DriverName) VALUES ("' + driverid + '", "' + car + '", ' + str(int(datetimeon.timestamp())) + ', ' +str(int(datetimeoff.timestamp())) + ', ' + str(shifttype) + ', ' + str(datetimeon.year) + str(datetimeon.month).zfill(2) + str(datetimeon.day).zfill(2) + ', ' + str(shiftlength.seconds) + ', "' + str(drivername) + '");')
print("Shifts CSV Okay!")
#Write Log file
servicelogname = 'serviceLog' + userInput.logonPath.replace('.', '_').replace('/', '_') + '.csv'
servicelog = open(servicelogname, 'w')
servicelog.write(serviceHours)
servicelog.close()
#get the start and end dates
c.execute('select min(Start) from Shifts;')
startdate = c.fetchone()
c.execute('select max(Start) from Shifts;')
enddate = c.fetchone()
daterange = str(datetime.fromtimestamp(startdate[0])) + ' to ' + str(datetime.fromtimestamp(enddate[0]))
print(daterange)

#instert charges into db
chargenumber = 0
with open(userInput.chargePath) as charges:
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

#car types 0 is sedan 1 is van

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
    "23":1,
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
c.execute('select distinct DriverID from Shifts')
driverrows = []
shifnums = 0
ownernums = 0
for row in c.fetchall():
    driverrows.append(row[0])

#get all the cars in range
c.execute('select distinct Car from Shifts')
carrows = {}
for row in c.fetchall():
    carrows[row[0].lstrip("V")] = []

#Open Template
templatestring = ""
templatefile = open('templates/statement-template.html', 'r')
templatestring = templatefile.read()

#Open car Template
cartemplatestring = ""
cartemplatefile = open('templates/car-template.html', 'r')
cartemplatestring = cartemplatefile.read()

driverPdfIndex = 0
totalDriverPdf = len(driverrows)
print("Writing Operator statements, this may take some time..")
#collapse shifts for each driver
for driver in driverrows:
    #c.execute('select Charge, Credit from Charges where DriverID = ' + str(driver) + ';')
    c.execute('select ' + tablecols + ' from Extras where ID = ' + str(driver) + ';')
    chargerow = c.fetchone()
    rowoutput = '<tr><td>' + str('CAR').ljust(10).replace(' ', '&nbsp;') + str('TYPE').ljust(10).replace(' ', '&nbsp;') + str('DATE').ljust(15).replace(' ', '&nbsp;') + str('SHIFT').ljust(12).replace(' ', '&nbsp;') + '</td><td></td></tr>'
    totalvalue = float(0.00)
    totalchargevalue = float(0.00)
    chargeoutput = ""
    if chargerow != None:
        for cnt, chargeval in enumerate(chargerow):
            if cnt > 2 and chargeval != '':
                totalchargevalue += chargeval
                chargeoutput += '<tr><td>' + chargeheader[cnt] + '</td><td>' + formatCharges(chargeval) + '</td></tr>'
        balancein = chargerow[1]
        balanceout = chargerow[2]
        chargeoutput += '<tr><td>Previous Account Balance</td><td>' + formatCharges(balancein) + '</td></tr>'
        #add balance forward to total
        totalchargevalue += balancein
    else:
        balancein = 0
        balanceout = 0
    driverID = ""
    drivername = ""
    thistemplate = ""
    

    

    for date in daterows:
        c.execute('select Car, Start, End, Type, Length, DriverName from Shifts where Date = ' + str(date) + ' and DriverID = ' + str(driver) +';')
        shifttime = 0
        startdates = []
        enddates = []
        types = []
        shiftvalue = 0
        driversname = []
        shiftType = ""
        cartypestring = ""
        carList = {}
        theseResults = c.fetchall()
        for row in theseResults or []:
            if row[3] < 2:
                types.append(row[3])
                shifttime += row[4]
                startdates.append(row[1])
                driversname.append(row[5])
                enddates.append(row[2])
                thiscar = row[0].lstrip("V")
                if thiscar in carList:
                    carList[thiscar]['length'] = row[4] + carList[thiscar]['length']
                else:
                    carList[thiscar] = {'length' : row[4], 'type' : row[3]}
            elif row[3] == 2:
                if round(((row[4]/60)/60),2) > 4:
                    print('--Time Loss Warning-- ' + str(round(((row[4]/60)/60),2)) + 'h ' + str(datetime.fromtimestamp(row[1])) + ' to ' + str(datetime.fromtimestamp(row[2])) + ' ' + str(row[5]) + ' ' + str(driver))
                
            elif row[3] == 3:
                if round(((row[4]/60)/60),2) > 4:
                    print('--Time Loss Warning-- ' + str(round(((row[4]/60)/60),2)) + 'h ' + str(datetime.fromtimestamp(row[1])) + ' to ' + str(datetime.fromtimestamp(row[2])) + ' ' + str(row[5]) + ' ' + str(driver))
                
            else:
                print('shift skipped, lost ' + str(round(((row[4]/60)/60),2)) + 'h')
        for carname in carList.keys():
            cartime = carList[carname]['length']
            #do write out the days data    
            if cartime > 0:
                weekday = str(datetime.fromtimestamp(startdates[0]).weekday())
                thiscar = carname
                #figure out the lease value
                if thiscar not in cartypes:
                    shiftType = 'ERROR'
                    print("car type error!")
                elif cartypes[thiscar] == 0:
                    #car
                    cartypestring = "C"
                    if carList[carname]['type'] == 0:
                        shiftType = 'DAY'
                        shiftvalue = leaseValues.getSedanRate('day')
                    elif carList[carname]['type'] == 1:
                        shiftType = 'NIGHT'
                        if userInput.longWeekendDateTime != "none" and datetime.fromtimestamp(startdates[0]).date() == userInput.longWeekendDateTime.date():
                            shiftvalue = 120 #check if longweekend
                        else:
                            shiftvalue = leaseValues.getSedanRate(weekday)
                    else:
                        print("shift type error!")
                elif cartypes[thiscar] == 1:
                    #van
                    cartypestring = "V"
                    if carList[carname]['type'] == 0:
                        shiftType = 'DAY'
                        shiftvalue = leaseValues.getVanRate('day')
                    elif carList[carname]['type'] == 1:
                        shiftType = 'NIGHT'
                        if userInput.longWeekendDateTime != "none" and datetime.fromtimestamp(startdates[0]).date() == userInput.longWeekendDateTime.date():
                            shiftvalue = 120 #check if longweekend
                        else:
                            shiftvalue = leaseValues.getVanRate(weekday)
                    else:
                        print("shift type error!")
                elif cartypes[thiscar] == 2:
                    #tcar
                    cartypestring = "T"
                    shiftType = 'TCAR'
                    shiftvalue = leaseValues.getTcarRate()
                else:
                    shiftType = 'ERROR'
                    print("car type error!")

                #check if the owner is driving
                c.execute('select * from Exemptions where car = ' + str(thiscar) + ' AND driverID = ' + str(driver) +';')
                thisOnwer = c.fetchall()
                if len(thisOnwer) != 0:
                    shiftvalue = 0
                    ownernums += 1
                rowoutput = rowoutput + '<tr><td>VT' + str(thiscar).ljust(8).replace(' ', '&nbsp;') + str(cartypestring).ljust(10).replace(' ', '&nbsp;') + str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7]).ljust(15).replace(' ', '&nbsp;') + allweekdays[weekday].ljust(12).replace(' ', '&nbsp;') + shiftType.ljust(8).replace(' ', '&nbsp;') + '</td>' + '<td class="rightalign">$' + str(shiftvalue) + '.00</td></tr>'
                carrows[thiscar].append({"driver" : str(driver), "car" : str(thiscar), "shift" : str(shiftType), "date" : str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7]), "value" : str(shiftvalue), "name" : driversname[0]})
                driverID = str(driver)
                drivername = driversname[0]
                totalvalue = totalvalue + int(shiftvalue)
                if int(shiftvalue) > 0:
                    shiftReport.appendShift(driver,str(thiscar),str(date),carList[carname]['type'],int(shiftvalue), drivername)#add shift to csv
                shifnums +=1

    #add values to template
    thistemplate = templatestring
    thistemplate = thistemplate.replace('$DATERANGE', daterange)
    thistemplate = thistemplate.replace('$STATEMENT_NOTE', str(userInput.statementNote))
    thistemplate = thistemplate.replace('$DRIVER_NAME', drivername)
    thistemplate = thistemplate.replace('$DRIVER_NUMBER', driverID)
    thistemplate = thistemplate.replace('$SHIFT_DATA', rowoutput)
    thistemplate = thistemplate.replace('$CHARGE_ROWS', chargeoutput)
    thistemplate = thistemplate.replace('$LEASE_TOTAL', str(Decimal(totalvalue).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)))
    thistemplate = thistemplate.replace('$TOTAL_BALANCE', str(Decimal(totalchargevalue - totalvalue).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    thistemplate = thistemplate.replace('$TO_DRIVER1', str((Decimal(totalchargevalue - totalvalue - balanceout)/2).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    thistemplate = thistemplate.replace('$TO_DRIVER2', str(Decimal((totalchargevalue - totalvalue - balanceout)/2).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    thistemplate = thistemplate.replace('$ACCOUNT_BALANCE', str(Decimal(balanceout).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    if balanceout > balancein:
        thistemplate = thistemplate.replace('$ACCOUNT_NOTES','Your minimum account balance (deposit) has been increased by $' + str(Decimal(balanceout - balancein).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;')
    else:
        thistemplate = thistemplate.replace('$ACCOUNT_NOTES',' ')
    thistemplatefilename = 'statements/operator/thisdriver.html'
    thistemplatefilenamepdf = 'statements/operator/' + str(driver) + '_' + drivername.replace(' ', '_').replace('/', '_') + '_statement_' + userInput.logonPath.replace('.', '_') + '.pdf'
    thistemplatefile = open(thistemplatefilename, 'w')
    thistemplatefile.write(thistemplate)
    thistemplatefile.close()
    #write pdf version of operator statement
    pdfkit.from_file(thistemplatefilename, thistemplatefilenamepdf,options = options, configuration = config)
    print('Writing Operator Statement ' + str(driverPdfIndex) + ' of ' + str(totalDriverPdf))
    driverPdfIndex += 1
    
driverfile = 'driver_' + userInput.logonPath.replace('.', '_').replace('/', '_') + '.csv'
shiftReport.writeFile(driverfile)
print(str(shifnums) + " shifts written")
templatefile.close()

c.close()
memconn.close()
print("Writing Car statements, this may take some time..")
#write day car shifts
genStatement(carrows, "DAY", 'templates/car-template.html')
genStatement(carrows, "NIGHT", 'templates/car-template.html')

print("all done!")
print('Press enter to close this window')
file = input()
