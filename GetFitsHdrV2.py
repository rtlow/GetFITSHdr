#required packages
import os
import re
import glob
import astropy.io.fits as fts
import pandas as pd
import copy


#receiving the data directory

cwd = os.getcwd()

print('Current Working Directory is:')
print(cwd)

#declaring these variables globally so that they can be used between blocks

#response for the first question
resp1 = 'a'

#response for the second question
resp2 = 'b'

#the directory that the user inputs
drty = 'c'

#the data directory path
ipath = 'd'

#the data directory path with '*.fits' appended; used for glob
dpath = 'e'

#getting either data directory or subdirectory from user
while True:
    print('Is the data in a subdirectory? (y/n)')
    resp1 = input()

    if resp1 == 'y' or resp1 == 'Y' or resp1 == 'n' or resp1 == 'N':
        break

#tried to make the process as foolproof as possible; adds forward slashes for you
while True:
    if resp1 == 'y' or resp1 == 'Y':
        print("Please input the data subdirectory")
        print("i.e. /data/")
        drty = input()

        if not (drty.endswith('/')):
            drty = drty + '/'

        if not (drty.startswith('/')):
            drty = '/' + drty

        ipath = cwd + drty

        print("The data path is")
        print(ipath)

        print("Is this the correct path? (y/n)")
        resp2 = input()

        if resp2 == 'y' or resp2 == 'Y':
            break

    elif resp1 == 'n' or resp1 == 'N':
        print("Please input the full data directory")
        print("i.e. /home/users/shared/spex/data/")

        drty = input()

        if not (drty.endswith('/')):
            drty = drty + '/'

        if not (drty.startswith('/')):
            drty = '/' + drty

        ipath = drty

        print("The data path is")
        print(ipath)
        
        print("Is this the correct path? (y/n)")
        resp2 = input()

        if resp2 == 'y' or resp2 == 'Y':
            break


#full filepath adds the file extension for use with glob
path = ipath + '*.fits'

#list containing data types to be collected from the header
#adding more data types should be done at the end of the list
hdtype = ['OBJECT', 'TCS_RA', 'TCS_DEC', 'DATATYPE', 'PROG_ID', 'OBSERVER', 'DATE_OBS', 'TIME_OBS', 'ITIME', 'POSANGLE', 'SLIT', 'GRAT', 'GFLT', 'TCS_HA', 'TCS_PA', 'TCS_AM']

#a different copy is used for the names of the excel columns which has the filenames as the first element in the list
hdtypes = copy.deepcopy(hdtype)

hdtypes = ['FNAME'] + hdtypes

#empty lists for containing the collected data
FNAME = []

#file paths are stored separately from filenames, since astropy needs the full paths to read the files
FPATH = []

OBJECT = []
TCS_RA = []
TCS_DEC = []
DATATYPE = []
PROG_ID = []
OBSERVER = []
DATE_OBS = []
TIME_OBS = []
ITIME = []
POSANGLE = []
SLIT = []
GRAT = []
GFLT = []
TCS_HA = []
TCS_PA = []
TCS_AM = []

#list of the data lists
#adding more data types should be done at the end of the list, but before "FNAME"
head = [OBJECT, TCS_RA, TCS_DEC, DATATYPE, PROG_ID, OBSERVER, DATE_OBS, TIME_OBS, ITIME, POSANGLE, SLIT, GRAT, GFLT, TCS_HA, TCS_PA, TCS_AM, FNAME]

#adding the filepaths to the FPATH list
for filename in glob.glob(path):
    FPATH.append(filename)

#places the paths in alphabetical order
FPATH.sort()

#filenames are extracted from the filepaths and added to the FNAME list
for filename in FPATH:
    fplen = len(ipath)
    head[len(hdtype)].append(filename[fplen:])

#for each file, extract the header data
for filename in FPATH:
    hdul = fts.open(filename)
    try:

        #for each data type, put it in its corresponding list in the list of lists 'head'
        for i in range(len(hdtype)):
            head[i].append(hdul[0].header[hdtype[i]])

    #if there is an error in reading the header data, print the offender's filename
    except KeyError:
        print("Error finding header data in: " + filename[fplen:])

#empty dict is used to store the datatype and the corresponding list
d = {}

#populate the dict using the dtype string as the key and the data list as the value
for i in range(len(head)-1):
    d[hdtype[i]] = head[i]

#add a final key data pair for the filenames
d[hdtypes[0]] = head[len(head)-1]

#create the excel file using the dict and the list of column names
df = pd.DataFrame(data=d, columns=hdtypes)

#get the name of the excel file
print("Please name your file:")
print("e.g. 2016B056.xlsx")
xlname = input()

#tried to make it foolproof; puts the file extension on for you
if not (xlname.endswith('.xlsx')):
    xlname = xlname + '.xlsx'

writer = pd.ExcelWriter(xlname, engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1')

writer.save()
