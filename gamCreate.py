# ~ import xlrd # to convert numbers to column letters
# ~ from xlsxwriter.utility import xl_rowcol_to_cell
# ~ import xlwt # to convert numbers to column letters, zero indexed
# import pyodbc # M$Sql
# ~ from openpyxl import Workbook # row/column 1/1 = A1
# ~ from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, NamedStyle
from openpyxl import load_workbook # read the files

# ~ import openpyxl
# ~ import re
# ~ import sys
import glob
import os #  for pipe and environ
import csv
import subprocess

# "Code" file info - which ones to process
codesfilename = "codes.xlsx"
school2ou = {}
school2notes = {}
school2email = {}
school2location = {}
school2destiny = {}
s2dcsv = ['SCHOOL,BARCODE,CF1:NAME,CF2:NOTES']

# from Connie's file(s), s/n to tag# and school
sntotag = {}
sntoschool = {}
cmd = []


				
def doesthisouexistingoogleadmin(ou):
	with os.popen('gam info org "{}" nousers'.format(ou)) as pipe:
		a = pipe.readlines()
		lena = len(a)
		if (0 == lena):
			# ~ print("No good, should I check up the tree?")
			return False
		else:
			return True

def gamcreatecommand(ou):
	return 'gam create org "{}"'.format(ou)

def createou(ou):
	# if creating OU has no /, then it's at the top level, and probably NOT right
	# if creating OU, find a good base, then build all the OUs after that base
	if not isinstance(ou,str):
		print("ERROR, I only do one OU at a time, something is terribly wrong.")
		exit()
	if ou[-1] == "/":
		print("ERROR, OUs should not end with trailing slash, fix excel and try again.")
		exit()
	if ('/' in ou.lstrip('/')):
		b = '/'.join((ou.split('/'))[:-1])
		if not (doesthisouexistingoogleadmin(b)):
			r = []
			r.append(gamcreatecommand(ou))
			r.extend(createou(b)) # careful, careful, ...
			return r
		else:
			r = []
			r.append(gamcreatecommand(ou))
			return r
	else:
		if (doesthisouexistingoogleadmin(ou)):
			return []
		else:
			print("ERROR, no new root OU created {}, I just won't do it, I just can't bring myself to belive this is correct.  If it is, please create the root manually at admin.google.com, then try again.".format(ou))
			exit()

# work from current directory, open each excel file
for f in glob.glob("*.xls*"):
	# skip any temp files
	if f[0] == "~":
		continue
	if f == codesfilename:
		print ("Gathering which school(s) to process, from {}. Error 404's are OK.".format(f))
		wb = load_workbook(f, read_only=True, data_only=True)
		ws1 = wb.active
		for r in range(2,30):
			school = ws1["a{}".format(r)].value
			targetou = ws1["b{}".format(r)].value
			emailconf = ws1["c{}".format(r)].value
			notes = ws1["d{}".format(r)].value
			location = ws1["e{}".format(r)].value
			destiny = ws1["f{}".format(r)].value
			if (isinstance(emailconf,str) and isinstance(targetou,str) and isinstance(school,str) and school[:2].upper() != "EX"):
				# ~ print ("s/n: ;;{};; = tag ;;{}CB-{};;".format(serial,school,tag))
				school = school.strip()
				targetou = targetou.strip()
				emailconf = emailconf.strip()
				if (isinstance(notes,str) ):
					school2notes[school] = notes.strip()
				else:
					school2notes[school] = ""
				if (isinstance(location,str) ):
					school2location[school] = location.strip()
				else:
					school2location[school] = ""
				
				destinypartofmessage = ""
				if (destiny == "Yes"):
					school2destiny[school] =  True
					destinypartofmessage = " and added to Destiny file"
				school2ou[school] = targetou
				school2email[school] = emailconf
				ougood = doesthisouexistingoogleadmin(targetou)
				if ougood:
					newouornot = ""
				else:
					newouornot = "a NEW ou "
					# check up one (and up one...) until you find a good root, build all needed OUs
					cmd.extend(createou(targetou))
				print ("*** {} Chromebooks in / by {} will be processed into {}{}{}.".format(school, emailconf, newouornot, targetou, destinypartofmessage))
		wb.close()
	else:
		print ("Adding lookups for serial/tags from {}".format(f))
		wb = load_workbook(f, read_only=True)
		for s in wb.sheetnames:
			# ~ print("worksheet {}".format(s))
			ws1 = wb[s]
			for r in range(14,44):
				tag = ws1["b{}".format(r)].value
				serial = ws1["c{}".format(r)].value
				desc = ws1["d{}".format(r)].value
				school = ws1["e{}".format(r)].value
				room = ws1["f{}".format(r)].value
				fulltag = "{}CB-{}".format(school,tag)
				if (isinstance(serial,str) and serial not in ("No Number","None") and isinstance(school,str)):
					# ~ print ("s/n: ;;{};; = tag ;;{}CB-{};;".format(serial,school,tag))
					# ~ print(ws1['A18'].value)
					serial = serial.strip()
					school = school.strip()
					if serial in sntotag.keys():
						# just ignore if same, otherwise red flag
						if sntotag[serial] != fulltag:
							print ("ERROR, dup serial: {} <> {}".format(sntotag[serial],fulltag))
							# probably do something else here, this could be a real problem TODO
							sntotag[serial] = 'dups {} {}'.format(sntotag[serial].replace("dups ",""),fulltag)
					else:
						sntotag[serial] = fulltag
						sntoschool[serial] = school
		wb.close()
	# ~ for a in sntotag.keys():
		# ~ print ("key {} = {}".format(a,sntotag[a]))

# if I do need to create new OUs, I only need to do each base once, remove dups
cmd = sorted(list(set(cmd)))

# ~ print("Ok, let's look at those s/n tags I found: {}".format(sntotag))
# ~ exit()

# ~ Now loop over all the OU / CrOS devices and add them to the list, if appropriate
with os.popen('gam print cros limit_to_ou / fields deviceId,serialNumber,status,lastSync,annotatedUser,annotatedLocation,annotatedAssetId,lastEnrollmentTime,orgUnitPath,notes') as pipe:
	reader = csv.DictReader(pipe)
	for row in reader:
		# ~ if this chromebook sn -> connie's file for school and tag -> codes for owners email
		# ~ If it is a good one to register, do it
		sn = row['serialNumber']
		cbuser = row['annotatedUser']
		status = row['status']
		if (status == "ACTIVE" and sn in sntoschool.keys() and sntoschool[sn] in school2email.keys() and school2email[sntoschool[sn]] == cbuser):
			school = sntoschool[sn]
			deviceid = row['deviceId']
			loc = school2location[school]
			mname = sntotag[sn]
			notes = school2notes[school]
			destou = school2ou[school]		
			cmd.append('gam update cros query:id:{} notes "{}" ou "{}" assetid "{}" location "{}"'.format(sn,notes,destou,mname,loc))
			# print(cmd[-1])
			if (school2destiny[school]):
				# add to destiny CSV also
				s2dcsv.append("{},{},{},{}".format(school,sn,mname,notes))

print("--- Here's what I'm really going to do if you say yes: ---")
for a in cmd:
	print (a)

l = len(cmd)
if l > 0:
	sifplural = ""
	if l > 1:
		sifplural = "s"
	doit = input("You sure you want to run {} command{} [Y,n]? ".format(len(cmd),sifplural))
	if (doit == "" or doit.lower()[0] == "y"):
		print ("Ok, this will just take a minute...")
		for c in cmd:
			# execute c and display any errors to screen
			print(c)
			r = subprocess.run(c, capture_output=True)
			if (r.stderr):
				print ("  ERROR -- {}".format(r.stderr))
		# create Destiny CSV if s2dcsv > 1 value
		if len(s2dcsv) > 1:
			user = os.environ['USERNAME']
			filename = "{}DestinyChromebooksFile.csv".format(user)
			# drop header line if appending to the file
			if os.path.exists(filename):
				del s2dcsv[0]
			f = open(filename, "a")  
			for l in s2dcsv:
				f.write("{}\n".format(l))
			f.close()
			print ("filename {} created/appended for Destiny.".format(filename))
	else:
		print("Maybe next time, thanks.")
else:
	print("Sorry, but I couldn't find anything to do.")



