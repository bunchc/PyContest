#!/usr/bin/python -tt

# PyContest.py
# Reads in a Google Spreadsheet and based on selected columns, builds a list of users.
# From each list, selects a winner.
# 
# With special thanks to Matthew Brender 
# 11.20.12 

import sys, string, re
import base64, gspread
import collections
from random import choice

# Change these!
user = 'user@gmail.com'
passwd = 'password'

# A list of every winner, to ensure no duplicates
winnerList = []

# Remember: URL is to the spreadsheet
# Example: spreadsheet = gc.open_by_url("https://docs.google.com/spreadsheet/ccc?key=1AkgN6IO_5rdrprh0V1pQVjFlQ2mIUDd3VTZ3ZjJubWc#gid=0")
url = "https://docs.google.com/path/to/doc"


############ ssLogin ###############
# Logs in and opens spreadsheet
# Returns spreadsheet
##
def ssLogin(url): 
    password = passwd #base64.b64decode(passwd)
    gc = gspread.login(user, password)
    spreadsheet = gc.open_by_url(url)
    print "successfully opened"
    return spreadsheet

######### markDupes #########
# Takes in the spreadsheet, gathers all data, and makes an attempt to clean obvious duplicates from the list
# It's not very effective due to the nature of the data in the list of dicts however...
##
def markDupes(spreadsheet): 
    worksheet = spreadsheet.sheet1
    allData = worksheet.get_all_records()
    newList = []
    for e in allData:
    	if e not in newList:
    		newList.append(e)
    return(newList)

######### listMaker #########
# Takes in a clean list of dict and a key.
# Returns a list of all things matching said key.
##
def listMaker(allEntries, prizeKey):
	count = 0
	prizeList = []
	for i in allEntries:
		if i[prizeKey] is not "":
			prizeList.append(i)
	return(prizeList)

######### winnerChooser #########
# Takes in a list of dict, selects a random list position, and returns the dict of the winner.
# Also makes sure the winner has not already been selected.
##
def winnerChooser(contestList):
	inList = False
	while inList is False:
		winner = choice(contestList)
		if winner not in winnerList:
			winnerList.append(winner)
			inList = True
		else:
			print "picking again\n"
	return(winner)

######### pearsonWinners #########
# Takes in a list of dict for Pearson, then chooses 10 at random
# finally, updates the spreadsheet with the winners.
# There are a few like it, I likely could have been a bit more generic here
##
def pearsonWinners(spreadsheet, pearsonList):
	i = 0
	pearsonWinner = []
	while i < 10:
		i += 1
		pearsonWinner.append(winnerChooser(pearsonList))
	updateSheet(spreadsheet, 'Pearson', pearsonWinner)

def trainSignal(spreadsheet, contestList):
	i = 0
	trainSignalWinner = []
	while i < 2:
		i += 1
		trainSignalWinner.append(winnerChooser(contestList))
	updateSheet(spreadsheet, 'TrainSignal', trainSignalWinner)

def sybexWinners(spreadsheet, contestList):
	i = 0
	sybexWinners = []
	while i < 7:
		i += 1
		sybexWinners.append(winnerChooser(contestList))
	updateSheet(spreadsheet, 'Sybex', sybexWinners)

######### updateSheet #########
# Takes in the spreadsheet, contest, & winner, adds a worksheet to our master workbook with the winner
##
def updateSheet(spreadsheet, contest, winner):
	newSheet = spreadsheet.add_worksheet(contest,1,1)
	# If a single contest, just update the sheet. Otherwise, add all the winners.
	if type(winner) is list:
		winnerKeys = winner[-1].keys()
		newSheet.append_row(winnerKeys)
		for w in winner:
			newSheet.append_row(w.values())
	else:
		newSheet.append_row(winner.keys())
		newSheet.append_row(winner.values())

####### main ########
def main():
	# Open and clean the spreadsheet
	spreadsheet = ssLogin(url)
	cleanList = markDupes(spreadsheet)
	# Carve the list down some
	veeamList = listMaker(cleanList, 'Veeam')
	vmturboList = listMaker(cleanList, 'VMTurbo')
	emcList = listMaker(cleanList, 'EMC')
	sybexList = listMaker(cleanList, 'Sybex')
	joshList = listMaker(cleanList, 'JoshAtwell')
	scottList = listMaker(cleanList, 'ScottLowe')
	trainsignalList = listMaker(cleanList, 'TrainSignal')
	pearsonList = listMaker(cleanList, 'VMwarePress')
	nutanixList = listMaker(cleanList, 'Nutanix')
	zertoList = listMaker(cleanList, 'Zerto')
	vmwareList = listMaker(cleanList, 'VMware')

	# Choose winner & update sheet for the easy contests
	sybexWinners(spreadsheet, sybexList)
	pearsonWinners(spreadsheet, pearsonList)
	trainSignal(spreadsheet, trainsignalList)
	updateSheet(spreadsheet, 'Veeam', winnerChooser(veeamList))
	updateSheet(spreadsheet, 'VMTurbo', winnerChooser(vmturboList))
	updateSheet(spreadsheet, 'EMC', winnerChooser(emcList))
	updateSheet(spreadsheet, 'Josh', winnerChooser(joshList))
	updateSheet(spreadsheet, 'Scott', winnerChooser(scottList))
	updateSheet(spreadsheet, 'Nutanix', winnerChooser(nutanixList))
	updateSheet(spreadsheet, 'Zerto', winnerChooser(zertoList))
	updateSheet(spreadsheet, 'VMware', winnerChooser(vmwareList))

if __name__ == '__main__':
	main()