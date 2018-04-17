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

def formatCharges(chargeval):
    if chargeval < 0:
        return '($' + str(Decimal(chargeval*int(-1)).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + ')'
    else:
        return '$' + str(Decimal(chargeval).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)) + '&nbsp;'

path_wkthmltopdf = r'wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
options = {
    'quiet': ''
    }
diskconn = sqlite3.connect('config/statements.db')
memconn = sqlite3.connect(':memory:')
query = "".join(line for line in diskconn.iterdump())
memconn.executescript(query)
c = memconn.cursor()
file=""
while file == "":
    print('Enter Report File Name (including file type "example.csv"): ')
    file = input()
    if file == "":
        print("No file selected")
chargesfile=""
while chargesfile == "":
    print('Enter Charges File Name (including file type "example.csv"): ')
    chargesfile = input()
    if chargesfile == "":
        print("No file selected")
print('Enter Statement Note/Reason or ENTER to skip: ')
statementnote = input()
if statementnote == "":
    statementnote = "None"
    print("Note Skipped.")

idpattern = re.compile(r'(\d\d*)')

optiontree = ET.parse('config/options.xml')
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

#get car owners
ownerlist = {}
with open('config/owner_id.csv') as ownerids:
    #read line by line
    for cnt, line in enumerate(ownerids):
        owneridvals = line.split(",")
        if cnt == 0:
            print('Reading Owner IDs')
        else:
            ownerlist[
                str(owneridvals[0])
                ] = {
                    'day' : owneridvals[1],
                    'night' : owneridvals[3]}

#insert csv into db
with open(file) as shiftrows:
    for cnt, line in enumerate(shiftrows):
        thisshiftrow = line.split(r'","')
        car = thisshiftrow[6]
        timeonstring = thisshiftrow[8]
        driverstring = thisshiftrow[7]
        timeoffstring = thisshiftrow[9]
        drivername = re.sub(r'(\d\d* \- )', ' ', driverstring)
        if(timeonstring != '-'):
            datetimeon = datetime.strptime(timeonstring, '%d %b %Y, %H:%M')
            if(timeoffstring == '-'):
                datetimeoff = datetimeon + timedelta(hours=8)
                print('logoff out of range, set to ' + str(datetimeoff) + ' for ' + drivername)
            else:
                datetimeoff = datetime.strptime(timeoffstring, '%d %b %Y, %H:%M')
            driverid = idpattern.match(driverstring).group(0)
            houron = datetimeon.hour
            houroff = datetimeoff.hour
            shiftlength = datetimeoff - datetimeon
            if houron > houroff:
                shifttype = 1 #night
            elif houron <= 4 and houroff <= 6 and shiftlength.seconds > 7200:
                #shift to yesterday
                datetimeon = datetimeon - timedelta(days=1)
                datetimeoff = datetimeoff - timedelta(days=1)
                shifttype = 1 #night
            elif houron <= 6 and houroff <= 6:
                shifttype = 3 #toss
            elif houron <= 6 and houroff >= 6 and shiftlength.seconds > 7200:
                shifttype = 0 #day
            elif houron >= 6 and houroff <= 17 and shiftlength.seconds > 7200:
                shifttype = 0 #day
            elif shiftlength.seconds > 7200:
                shifttype = 1 #night
            else:
                shifttype = 2 #toss
            
            c.execute('INSERT INTO Shifts (DriverID, Car, Start, End, Type, Date, Length, DriverName) VALUES ("' + driverid + '", "' + car + '", ' + str(int(datetimeon.timestamp())) + ', ' +str(int(datetimeoff.timestamp())) + ', ' + str(shifttype) + ', ' + str(datetimeon.year) + str(datetimeon.month).zfill(2) + str(datetimeon.day).zfill(2) + ', ' + str(shiftlength.seconds) + ', "' + str(drivername) + '");')
print("Shifts CSV Okay!")

#get the start and end dates
c.execute('select min(Start) from Shifts;')
startdate = c.fetchone()
c.execute('select max(Start) from Shifts;')
enddate = c.fetchone()
daterange = str(datetime.fromtimestamp(startdate[0])) + ' to ' + str(datetime.fromtimestamp(enddate[0]))
print(daterange)

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
driverdata = open('csv/' + driverfile, 'w')

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
        output = ""
        shiftType = ""
        cartypestring = ""
        carList = {}
        
        for row in c.fetchall() or []:
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
                        shiftvalue = sedanddaylease
                    elif carList[carname]['type'] == 1:
                        shiftType = 'NIGHT'
                        shiftvalue = nightsedanleases[weekday]
                    else:
                        print("shift type error!")
                elif cartypes[thiscar] == 1:
                    #van
                    cartypestring = "V"
                    if carList[carname]['type'] == 0:
                        shiftType = 'DAY'
                        shiftvalue = vanldayease
                    elif carList[carname]['type'] == 1:
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
                #check if the owner is driving
                if thiscar not in ownerlist:
                    shiftvalue = 0
                    ownernums += 1
                elif carList[carname]['type'] == 0 and int(ownerlist[thiscar]['day']) == int(driver):
                    shiftvalue = 0
                    ownernums += 1
                elif carList[carname]['type'] == 1 and int(ownerlist[thiscar]['night']) == int(driver):
                    shiftvalue = 0
                    ownernums += 1
                output = str(driver) + ', ' + str(thiscar) + ', ' + str(carList[carname]['type']) +', '+ str(date) +', '+ str(round(((shifttime/60)/60), 2)) + ', ' + str(shiftvalue) + ', ' + driversname[0] + '\n'
                rowoutput = rowoutput + '<tr><td>VT' + str(thiscar).ljust(8).replace(' ', '&nbsp;') + str(cartypestring).ljust(10).replace(' ', '&nbsp;') + str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7]).ljust(15).replace(' ', '&nbsp;') + allweekdays[weekday].ljust(12).replace(' ', '&nbsp;') + shiftType.ljust(8).replace(' ', '&nbsp;') + '</td>' + '<td class="rightalign">$' + str(shiftvalue) + '.00</td></tr>'
                carrows[thiscar].append({"driver" : str(driver), "car" : str(thiscar), "shift" : str(shiftType), "date" : str(str(date)[0] + str(date)[1] + str(date)[2] +str(date)[3] + '/' + str(date)[4] +str(date)[5] + '/' + str(date)[6] +str(date)[7]), "value" : str(shiftvalue), "name" : driversname[0]})
                driverID = str(driver)
                drivername = driversname[0]
                totalvalue = totalvalue + int(shiftvalue)
                alldriverdata += output
                shifnums +=1

    #add values to template
    thistemplate = templatestring
    thistemplate = thistemplate.replace('$DATERANGE', daterange)
    thistemplate = thistemplate.replace('$STATEMENT_NOTE', str(statementnote))
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
    thistemplatefilename = 'statements/operator/' + str(driver) + '_' + drivername.replace(' ', '_').replace('/', '_') + '_statement_' + file.replace('.', '_') + '.html'
    thistemplatefilenamepdf = 'statements/operator/' + str(driver) + '_' + drivername.replace(' ', '_').replace('/', '_') + '_statement_' + file.replace('.', '_') + '.pdf'
    thistemplatefile = open(thistemplatefilename, 'w')
    thistemplatefile.write(thistemplate)
    thistemplatefile.close()
    #write pdf version of operator statement
    pdfkit.from_file(thistemplatefilename, thistemplatefilenamepdf,options = options, configuration = config)
    print('Writing Operator Statement ' + str(driverPdfIndex) + ' of ' + str(totalDriverPdf))
    driverPdfIndex += 1
    

driverdata.write(csvheader + alldriverdata)
print(str(shifnums) + " shifts written")
templatefile.close()
driverdata.close()
c.close()
memconn.close()
print("Writing Car statements, this may take some time..")
#write car statements
progress = 1
totalstatements = len(carrows.items())
for car, shifts in carrows.items():
    leasetotal = 0
    shiftoutput = '<tr><td>' + str('CAR').ljust(10).replace(' ', '&nbsp;') + str('DRIVER').ljust(30).replace(' ', '&nbsp;') + str('DATE').ljust(15).replace(' ', '&nbsp;') + str('SHIFT').ljust(12).replace(' ', '&nbsp;') + '</td><td></td></tr>'
    for shift in shifts:
        leasetotal += int(shift['value'])
        shiftoutput += '<tr><td>' + shift['car'].ljust(10).replace(' ', '&nbsp;') + shift['driver'].ljust(8).replace(' ', '&nbsp;') + shift['name'].ljust(22).replace(' ', '&nbsp;') + shift['date'].ljust(15).replace(' ', '&nbsp;') + shift['shift'].ljust(10).replace(' ', '&nbsp;') +'</td><td class="rightalign">$' + shift['value'] + '.00</td></tr>'
    carthistemplate = cartemplatestring
    carthistemplate = carthistemplate.replace('$DATERANGE', daterange)
    carthistemplate = carthistemplate.replace('$STATEMENT_NOTE', str(statementnote))
    carthistemplate = carthistemplate.replace('$CAR_NAME', car)
    carthistemplate = carthistemplate.replace('$SHIFT_DATA', shiftoutput)
    carthistemplate = carthistemplate.replace('$LEASE_TOTAL', str(Decimal(leasetotal).quantize(Decimal('0.01'), rounding=ROUND_HALF_DOWN)))
    carthistemplatefilename = 'statements/car/' + str(car) + '_statement_' + file.replace('.', '_') + '.html'
    carthistemplatefilenamepdf = 'statements/car/' + str(car) + '_statement_' + file.replace('.', '_') + '.pdf'
    carthistemplatefile = open(carthistemplatefilename, 'w')
    carthistemplatefile.write(carthistemplate)
    carthistemplatefile.close()
    
    pdfkit.from_file(carthistemplatefilename, carthistemplatefilenamepdf,options = options, configuration = config)
    print("Written car statement " + str(progress) + " of " + str(totalstatements))
    progress += 1

print("all done!")
print('Press enter to close this window')
file = input()
