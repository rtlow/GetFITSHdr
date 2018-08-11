#required packages
import os
import tkinter as tk
from tkinter import messagebox as mb
from tkinter import filedialog as fd
import glob
import astropy.io.fits as fts
import pandas as pd
import copy


#receiving the data directory

cwd = os.getcwd()

#create root for tk
root = tk.Tk()

#tk datatypes for holding the path info
inPath = tk.StringVar()

outPath = tk.StringVar()

#setting the window title
root.title("GetFITSHdr")


def getFitsHdr():
  
  ipath = inPath.get()
  
  opath = outPath.get()
  
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
          fileErrorMessage(filename[fplen:])

  #empty dict is used to store the datatype and the corresponding list
  d = {}

  #populate the dict using the dtype string as the key and the data list as the value
  for i in range(len(head)-1):
      d[hdtype[i]] = head[i]

  #add a final key data pair for the filenames
  d[hdtypes[0]] = head[len(head)-1]

  #create the excel file using the dict and the list of column names
  df = pd.DataFrame(data=d, columns=hdtypes)


  writer = pd.ExcelWriter(opath, engine='xlsxwriter')

  df.to_excel(writer, sheet_name='Sheet1')

  writer.save()
  
  successMessage()



#tk Message for the program name and main message
T = tk.Message(root, justify=tk.CENTER, text="GetFITSHdr\n Please Select the Data Directory\n")
T.config(font=('times', 12))
T.grid(row=0, column=1)

#tk Text for the Data Dir label
dataDText = tk.Label(root, height=1, width=10, text="Data Dir")
dataDText.grid(row=1, column=0)

#tk Text for the input Dir field
K = tk.Text(root, height=1, width = 80)
K.grid(row=1, column=1)


'''
Function getiDir() uses a "Browse" dialog to get the desired directory for
the input data
'''
def getiDir():
  #uses tk's filedialog to ask for a data directory
  inPath.set(fd.askdirectory(parent=root, initialdir=cwd, title='Please select a directory'))
  
  #checks if the path exists, if so continues
  if len(inPath.get()) > 0:
    if not(inPath.get().endswith('/')):
      inPath.set(inPath.get() + '/')
    
    K.delete(1.0, tk.END)
    K.insert(tk.END, inPath.get())

#button to activate the browse dialog
ibut = tk.Button(root, text="Browse", command=getiDir)
ibut.grid(row=1, column=2)

#tk text for the Out Dir label
outDText = tk.Label(root, height=1, width=10, text="Output Dir")
outDText.grid(row=2, column=0)

#tk text for the Out Dir field
M = tk.Text(root, height=1, width = 80)
M.grid(row=2, column=1)

'''
Function getiDir() uses a "Browse" dialog to get the desired directory for
the output file
'''
def getoDir():
  #uses tk's filedialog to ask for the save as directory
  outPath.set(fd.asksaveasfilename(parent=root, initialdir=cwd, title='Save As', filetypes = (("Microsoft Excel", "*.xlsx"), ("all files", "*.*"))))
  
  #checks if the path exists, if so continues
  if len(outPath.get()) > 0:
    if not(outPath.get().endswith('.xlsx')):
      outPath.set(outPath.get() + '.xlsx')
    
    M.delete(1.0, tk.END)
    M.insert(tk.END, outPath.get())


#button to activate the saveas dialog
obut = tk.Button(root, text="Browse", command=getoDir)
obut.grid(row=2, column=2)


'''
Function select() runs excelCollate() when the button is pressed and if the 
directories are not empty
'''
def select():
  if (len(inPath.get()) > 0)  and (len(outPath.get()) > 0):
     getFitsHdr()

 
def successMessage():
  #messagebox with success message
  mb.showinfo(title='Success!', message='Successfully wrote file at ' + outPath.get())

def fileErrorMessage( fname ):
  mb.showerror(title="Error", message="Error finding header data in: " + fname)
  
def errorMessage():
  mb.showerror(title="Error", message="Invalid File or Directory")
 
#button to activate the script 
sbut = tk.Button(root, text='Submit', command=select)
sbut.grid(row=3, column=1)

root.mainloop()