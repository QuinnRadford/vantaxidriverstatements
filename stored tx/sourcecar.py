print("enter the file")
thisfile = input()
output = "account,amount,comment\n"
print("enter the comment")
comment = input()
print("enter the source")
source = input()
print("enter the source acc")
sourceacc = input()
output += "06-06-18," + source + "," + comment + "\n"
with open(thisfile) as ownerids:
    #read line by line
    for cnt, line in enumerate(ownerids):
        owneridvals = line.split(",")
        if cnt == 0:
            print('Reading Owner IDs')
        else:
            output += owneridvals[1] + "," + owneridvals[2].replace('\n', "") + ",\n"
            output += sourceacc + ",-" + owneridvals[2].replace('\n', "") + ",\n"
filename = source + ".csv"
servicelog = open(filename, 'w')
servicelog.write(output)
servicelog.close()