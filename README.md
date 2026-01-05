# gamCreate4Chromebooks
Use some spreadsheets to feed Google's gam to create OUs and relocate provisioned chromebooks from /

Install Requirements

- gam  https://github.com/GAM-team/GAM/wiki/Downloads-Installs , follow all instructions to setup a few connections into Admin.Google.com.  If you get an Oauth error, run "gam oauth create".  When done, manually running "gam info domain" should return a few lines from your domain, not prompt you for things each time, repeat setup until then.
- python 3 https://www.python.org/downloads/ or Windows store, include it and pip3 in your PATH/Environment Variables
- py -m pip install openpyxl
- standard python3 library now includes glob, os, csv, subprocess)

- keep everything in the same directory

Running the Program

- copy codesExample.xlsx to codes.xls and edit to meet your needs
- Save inventory xls* file(s) in the same directory with code, School initials must be present in filename
- CMD in same directory as code, "python gamCreate.py", and hit enter when asked

gamCreate4Chromebooks.py logic

- loops over codes.xlsx for those with Emails, and builds School -> OU, Email, Notes and Description
- loops over all other Excel files mapping Serial -> Tag and School
- loops over / OU in admin.google and finds potential Chromebooks to move, mapping Serial -> email
- if found chromebook serial maps to valid inventory and codes entries, add it to the list to process
- build AnnotatedAssetID as SchoolCB-Tag for each Serial found
- review all "gam create org" and "gam update cros" commands proposed and ask user for approval
- do it if yes, including making/appending a CSV summary of changes

codes.xlsx file requirements

- copy from codesExample.xlsx for first use, or create from scratch
- Sheets, just one please, named whatever you like
- Rows, 2-29 are processed, search code for "range(2" if you need to change
- A: School(required) (initials, any starting with EX are ignored)
- B: Target OU(required) (new or existing OU to move chromebooks in to, ex: /ADMIN/Cart 14)
- C: AnnotatedUser(required) (this limits which chromebooks in / you want to process at this time)
- D: Notes(optional) (pushed into the Notes on each chromebook moved)
- E: Location(optional) (pushed into the Location on each chromebook moved)
- F: Prep for Destiny Import(y/n) (if y, processed chromebooks are appended to CSV file for import into library system, or sharing)

Other Excel xls* files to match Serial number to Inventory Tag number

- Sheets - use as many as you like
- Rows, 2-43 are processed, search code for "range(2" if you need to change
- B: Tag
- C: Serial
- E: School (matching codes.xlsx:A)

_tested only in Windows_
