#!/usr/bin/python3

from openpyxl import load_workbook # read the excel files
import glob # for looping over all Excel files
import os #  for pipe and environ
import csv # for processing output of gam print cros...
import subprocess # for running gam commands

import sys

def searchfile(needle, fnames):
	for f in glob.glob(fnames):
		wb = load_workbook(fname, read_only=True, data_only=True)
		for s in wb.sheetnames:
			ws1 = wb[s]
			for r in range(2,200):
				tag = ws1["b{}".format(r)].value
				serial = ws1["c{}".format(r)].value
				desc = ws1["d{}".format(r)].value
				school = ws1["e{}".format(r)].value
				room = ws1["f{}".format(r)].value
				if (needle in tag,serial,desc,school,room):
					print ("{},{},{},{},{}".format(tag,serial,desc,school,room))
		wb.close()

def getoptions(needle = sys.argv[1], fname = sys.argv[2]):
	for a in sys.argv:
		print ("found something: {}".format(a))
	# if needle > '' and fname > '':
	# 	return needle, fname
	# else:
	# 	exit # could be more graceful

if __name__ == "__main__":
	# excelgrep blah file.xlsx
	needle, fnames = getoptions()
	searchfile(needle, fnames)





