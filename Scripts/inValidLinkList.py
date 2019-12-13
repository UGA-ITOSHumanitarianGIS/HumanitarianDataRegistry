# -*- coding: utf-8 -*-
# --------------------------------------------------------------------------------------------------------------------------------
# inValidLinkList.py
# Created on: 2019-12-13 09:48:01.00000
#   
# Author: ADR
# Description: See Humanitarian IM Toolbox and UN contacts. Scans the Registry
#               of GIS applications and repositories listings as downloaded by a user here:
#               https://github.com/UGA-ITOSHumanitarianGIS/HumanitarianDataRegistry/raw/master/Data/GIS%20Data%20Repositories.xlsx
#               and reports those that are no longer accessible via a web request.
#               outputs the results for no access to noaccess.log file in current working
#               directory or the path where the spreadsheet resides. The 9th column
#               of the spreadsheet is anticipated to have the web addresses of the resources.
##
# Example: python inValidLinkList.py -xls-file "C:\Workspace\GIS & Data Platforms.xlsx"
#
# Called by: TBD
# --------------------------------------------------------------------------------------------------------------------------------


import os
import sys
import datetime
import urllib2
import requests
import xlrd
import argparse

refreshLog = os.path.join(os.getcwd(), "noaccess.log")#r"c$\\Workspace\\noaccess.log"
def log(message):
    with open(refreshLog, 'a') as logMessage:
        print('%s - %s\n' % (str(datetime.datetime.now()), str(message)))
        logMessage.write('%s - %s\n' % (str(datetime.datetime.now()), str(message)))    

    return

def doLinkCheck(linkAddress):

    try:

 # App log
        #The spreadsheet
        wb = xlrd.open_workbook(linkAddress)
        sheet = wb.sheet_by_index(0)
        print('Processing rows total: ' + str(sheet.nrows))

        #Now run through each entry and ping the address
        for i in range(sheet.nrows):
                try:
                    #print(sheet.cell_value(i, 9))
                    cURL = sheet.cell_value(i, 9)
                    if i <> 0:
                        ret = requests.get(cURL.strip())
                        if ret.status_code > 206:
                                log('Double check this for access!: ' + cURL)
			
                except:
                        log("Exception caught:  " + str(Exception.message))
                        pass
    except:
        log("Exception caught:  " + str(Exception.message))
        pass
    log("refresh completed!")
    return

# Run app with messaging and flow
log("Beginning python process...")
def make_path_sane(p):
        """Function to uniformly return a real, absolute filesystem path."""
        # ~/directory -> /home/user/directory
        p = os.path.expanduser(p)
        # A/.//B -> A/B
        p = os.path.normpath(p)
        # Resolve symbolic links
        p = os.path.realpath(p)
        # Ensure path is absolute
        p = os.path.abspath(p)
        refreshLog = os.path.join(os.path.dirname(p), "noaccess.log")
        return p

Parser = argparse.ArgumentParser(prog= "Parser")

Parser.add_argument("-xls-file", help= "Please indicate the path to the xls file.", required=True)

args = Parser.parse_args()
inFile = make_path_sane(args.xls_file)
doLinkCheck(inFile)
log("Python process complete...")
