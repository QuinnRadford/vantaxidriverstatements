def formatCarAccount(car, shift):
    shiftString = str(int(shift + 1)) #shift up shift type by 1
    carString = str(car)
    carString = carString.rjust(3,"0") #pad car number with 0s to the right
    carString = "22" + carString.ljust(5,"0") + shiftString
    return carString

output = ""
with open("cartypes.csv") as ownerids:
    #read line by line
    for cnt, line in enumerate(ownerids):
        owneridvals = line.split(",")
        if cnt == 0:
            print('Reading Owner IDs')
        else:
            output += owneridvals[0] + "," + formatCarAccount(owneridvals[0], 0) + "\n"
            output += owneridvals[0] + "," + formatCarAccount(owneridvals[0], 1) + "\n"
servicelog = open("caraccount.csv", 'w')
servicelog.write(output)
servicelog.close()