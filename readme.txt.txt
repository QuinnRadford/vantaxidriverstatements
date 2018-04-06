How to use the Statement Program:

Generating A Statement
1. Generate a “Car Log-on/Log-off Report” using the MTDATA reports tool under “Operational”
2. Export the report as CSV and save it to the same folder as the program
3. In excel or similar, create a sheet with the Driver ID column first. Add as many additional columns as you like. YOU MUST INCLUDE A HEADER COLUMN
4. Save the sheet as a CSV to the same folder as the program
5. Open the Statement Program statementcalculator.exe and follow the prompts, entering the file name of the xml and the csv
6. Enter a note if desired to be displayed on all statements
7. The statements will now be generated

Results:
Individual Operator and Car statemens in the “statements” folder.
csv export of all shifts worked in main program folder.
The names of the statements will be based on the name of the report used to generate them along with the Operator id and name.

Changing Lease Rates
1. Use Notepad to open “options.xml” in the config folder
2. The sedan rates are under the <sedan> tag
3. The van rates are under the <van> tag
4. The tcar rates are at the bottom under the <tcar> tag
5. Do not enter decimals, dollar signs, or anything other than numbers
6. Save when finished

Changing Owner-Operator information
1. Use excel or similar to open "owner_id.csv" in the config folder
2. make sure to enter the correct driver ID for the owner in the appropriate column (night or day)
3. The name/note columns do not affect the program, they are for human reference only