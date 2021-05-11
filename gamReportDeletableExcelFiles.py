#!/usr/bin/python3

from openpyxl import load_workbook # read the excel files
import glob # for looping over all Excel files
import os #  for pipe and environ
import csv # for processing output of gam print cros...
import subprocess # for running gam commands
 
codesfilename = "codes.xlsx"
school2ou = {}
school2notes = {}
school2email = {}
school2location = {}

gamexe = "gam"

# from Connie's file(s), s/n to tag# and school
# edit (14,44) if you use a different range of rows, around line  117
found = []
validschools = ('BES','BMS','MVM','DNE','MMS','EGP','ECE','NLE',\
	'MHE','MAE','MON','SUN','SMS','PES','GHS','UHS','ECHS','RHS','VVE','WAE')

# def schoolfromou(ou):
# 	parts = ou.split('/')
# 	for a in range(1,len(parts)+1):
# 		p = parts[-a]
# 		if p in validschools:
# 			return p
# 	return ""

def checkthisfile(f):
	if f[0] == "~" or f == codesfilename:
		return "Skipping file {}".format(f)
	else:
		print ("Checking on file {}".format(f))
		wb = load_workbook(f, read_only=True)
		for s in wb.sheetnames:
			# ~ print("worksheet {}".format(s))
			ws1 = wb[s]
			for r in range(2,44):
				serial = ws1["c{}".format(r)].value
				desc = ws1["d{}".format(r)].value
				school = ws1["e{}".format(r)].value
				if (isinstance(serial,str) and serial > "" and 'hromebook' in desc):
					serial = serial.strip()
					if serial not in found:
						return "Can't delete this one, I found {},{},{}".format(serial,desc,school)
		wb.close()
		return 'Ok to del "{}"'.format(f)
		
# ~ Loop over all the CrOS devices and add them to the list
# with os.popen('{} print cros fields serialNumber,annotatedAssetId,orgUnitPath'.format(gamexe)) as pipe:
with os.popen('{} print cros fields serialNumber,status,orgUnitPath'.format(gamexe)) as pipe:
	reader = csv.DictReader(pipe)
	for row in reader:
		# ~ if this chromebook sn -> connie's file for school and tag -> codes for owners email
		# ~ If it is a good one to register, do it
		status = row['status']
		ou = row['orgUnitPath']
		if (status == "ACTIVE" and ou != "/"):
			# ou = row['orgUnitPath']
			# school = schoolfromou(ou)
			# if (school):
			# 	deviceid = row['deviceId']
			# 	loc = school2location[school]
			# 	mname = sntotag[sn]
			found.append(row['serialNumber'])

# work from current directory, open each excel file
for f in glob.glob("*.xls*"):
	print(checkthisfile(f))

