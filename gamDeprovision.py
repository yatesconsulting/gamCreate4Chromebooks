#!/usr/bin/python3

from tkinter import E
from openpyxl import load_workbook # read the excel files
import glob # for looping over all Excel files
import os #  for pipe and environ
import csv # for processing output of gam print cros...
import subprocess # for running gam commands
import glob # for file globbing with extensions
import datetime

# debug = 1
deprovisionedOU  =  "/Deprovisioned"
deprovisionSoonOU = "/DeprovisionedSoon"
gamexe = "gam"
thismoment = datetime.datetime.today() 
cmd = []

def cbfromexcelfiles():
	# look through fixed asset retirement/delete excel files for chromebooks to delete
	files = glob.glob("*FARetirement*.xlsx")
	cb = {}
	for f in files:
		wb = load_workbook(f, read_only=True, data_only=True)
		ws1 = wb.active
		workorder = ws1["i19"].value
		print ("Gathering devices from {} workorder# {}.".format(f, workorder))
		for r in range(25,46):
			cbid = ws1["b{}".format(r)].value
			if cbid and cbid not in cb:
				if checkifchromebookexists(cbid):
					cb[cbid] = {'origin':f"{f}",'workorder':f"{workorder}"}
				else:
					print(f"{cbid} not found in admin.google.com, ignored")
			elif cbid:
				print(f"{cbid} found in multiple files or lines, {cb[cbid]['origin']} and {f}")
		wb.close()
	return cb

def cbfromOU():
	# ~ Loop over all the OU / CrOS devices and add them to the list, if appropriate
	cb = {}
	print ("Gathering devices from OU {}.".format(deprovisionSoonOU))
	with os.popen(f'{gamexe} print cros limit_to_ou {deprovisionSoonOU} fields annotatedAssetId') as pipe:
		reader = csv.DictReader(pipe)
		for row in reader:
			cbid = row['annotatedAssetId']
			if cbid:
				id = int(cbid.split('-')[-1])
				if id not in cb:
					cb[id] = {'origin':'OU'}
	return cb

def anotinb(a, b):
	cnt = 0
	for t in a:
		if t not in b:
			cnt += 1
	return cnt

def checkifgoodgaminuse(a):
	with os.popen(f'{gamexe} gam update cros "query:asset_id:{a}" updatenotes "#notes#') as pipe:
		reader = csv.DictReader(pipe)
		for row in reader:
			try:
				for row in reader:
					return True
			except:
				print("FATAL ERROR, gam must be upgraded to gamADV, see Bryan")
				return False
	return False # never here, I think

def checkifchromebookexists(a):
	with os.popen(f'{gamexe} info cros "query:asset_id:{a}" fields serialNumber') as pipe:
		reader = csv.DictReader(pipe)
		try:
			for row in reader:
				return True
		except:
			return False
	return False # never here, I think

def deprovisioncrosinv(a):
	cmd.append(f'{gamexe} issuecommand cros "query:asset_id:{a}" command remote_powerwash doit')
	cmd.append(f'{gamexe} issuecommand cros "query:asset_id:{a}" action deprovision_retiring_device acknowledge_device_touch_requirement')

def updatenote(a, note):
	cmd.append(f'{gamexe} update cros "query:asset_id:{a}" updatenotes "#notes#\\n{note}"')

def movetofinalourip(a):
	cmd.append(f'{gamexe} update cros "query:asset_id:{a}" ou "{deprovisionedOU}"')


def warnthenruncmd():
	print("--- Here's what I'm really going to do if you say yes: ---")
	# print ("*** {} Chromebooks in / by {} will be processed into {}{}{}.".format(school, emailconf, newouornot, targetou, destinypartofmessage))

	for a in cmd:
		print (a)

	l = len(cmd)
	if l > 0:
		sifplural = ""
		if l > 1:
			sifplural = "s"
		doit = input("You sure you want to run {} command{} [Y,n]? ".format(len(cmd),sifplural))
		if (doit.strip() == "" or doit.lower().strip()[0] == "y"):
			print ("Ok, this will take a few minutes...")
			for c in cmd:
				# execute c and display any errors to screen
				print(c)
				r = subprocess.run(c, capture_output=True)
				if (r.stderr):
					print ("  ERROR -- {}".format(r.stderr))
		else:
			print("Maybe next time, thanks.")
	else:
		print("Sorry, but I couldn't find anything to do.")

if __name__ == "__main__":
	# look for /DeprovisionedSoon or *FARetirement*.xls* files to process things from

	cb1 = cbfromexcelfiles()
	cb2 = cbfromOU()
	if cb1 or cb2:
		if (cb1 and checkifgoodgaminuse(cb1[0] or cb2[0])):
			for a in cb1:
				print(f"ph1: {a}")
				updatenote(a,f"work order {cb1[a]['workorder']}")
				updatenote(a,f"deprovisioned about {thismoment}")
				deprovisioncrosinv(a)
				movetofinalourip(a)
			for a in cb2:
				if a not in cb1:
					print(f"ph2: {a}")
					updatenote(a,f"deprovisioned about {thismoment}")
					deprovisioncrosinv(a)
					movetofinalourip(a)

			warnthenruncmd()
		else:
			print("Sorry, you need to upgrade your GAM version")
	else:
		print("I didn't find anything to do")