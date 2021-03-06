#Test to parse a Final Draft script into its xml elements to put
#into a Google Sheet for easier script breakdown and scheduling

import xml.etree.ElementTree as ET
import gspread
from gspread_formatting import *
from gspread_formatting import batch_updater
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

#hook up the sheet with the API
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("preproCREDS.json", scope)
client = gspread.authorize(creds)

#Enter the Final Draft script name as a .fdx
script = input('Script Name: ')

#import the data
tree = ET.parse(script)
root = tree.getroot()

#Function to get scene information and return them as a zipped list
#with scene number, scene heading, and page length

def sceneinfo():
	SceneNums = []
	Sluglines = []
	PageLength = []
	#get the scene number from the Scene Heading tag
	#got the slug line that is between the text tags at the top of the p
	for cont in root.findall('.Content/Paragraph'):
		if cont.attrib.get("Type") == "Scene Heading":
			SceneNums.append(cont.attrib.get("Number"))
			Sluglines.append(cont.find('Text').text)
	#get the scene length from the Scene Properties tag
	for cont in root.findall('.Content/Paragraph/SceneProperties'):
		PageLength.append(cont.attrib.get("Length"))
	#zip the lists to keep the appropriate numbers, headings, and length together
	#SceneFo = list(zip(SceneNums, PageLength, Sluglines))
	return list(zip(SceneNums, PageLength, Sluglines))

headings = ["Scene Number","Page Length", "INT/EXT", "Location", "D/N"]
topinfo = []
topinfo.append(headings)
topinfo.append(sceneinfo())

def scenelists():
	sheethead=[]
	for i in range(len(sceneinfo())):
		sheethead.append(topinfo[0])
		sheethead.append(topinfo[1][i])
	return sheethead

# this splits up the slugline into its parts
innerlist = []
for i in range(len(scenelists())):
	if i % 2 != 0:
		#Scene Number
		innerlist.append(scenelists()[i][0])
		# Page Length
		innerlist.append(scenelists()[i][1])
		# IorE = 
		innerlist.append(scenelists()[i][2].split('.')[0])
		# place = 
		innerlist.append(scenelists()[i][2].split('.')[1].split('-')[0])
		# DorN = 
		innerlist.append(scenelists()[i][2].split('.')[1].split('-')[1])

#to make a list of lists of lists for import to the Google API, I just appended a whole bunch
#of lists over and over. This is not elegant. 
ITAS1 = []
ITAS2 = []
ITAS3 = []
ITAS4 = []
ITAS5 = []
ITAS6 = []
ITAS7 = []

#this is hard coded for 7 scenes. Try to figure out a way to do this no
#matter how many scenes there are. There will always be 5 parts to the list
for i in range(len(innerlist)):
	if i < 5:
		ITAS1.append(innerlist[i])
	elif i < 10:
		ITAS2.append(innerlist[i])
	elif i < 15:
		ITAS3.append(innerlist[i])
	elif i < 20:
		ITAS4.append(innerlist[i])
	elif i < 25:
		ITAS5.append(innerlist[i])
	elif i < 30:
		ITAS6.append(innerlist[i])
	elif i < 35:
		ITAS7.append(innerlist[i])

#these are the lists of lists with the headings and the scene info in them
IMATOP1 = []
IMATOP2 = []
IMATOP3 = []
IMATOP4 = []
IMATOP5 = []
IMATOP6 = []
IMATOP7 = []
IMATOP1.append(scenelists()[0])
IMATOP1.append(ITAS1)
IMATOP2.append(scenelists()[0])
IMATOP2.append(ITAS2)
IMATOP3.append(scenelists()[0])
IMATOP3.append(ITAS3)
IMATOP4.append(scenelists()[0])
IMATOP4.append(ITAS4)
IMATOP5.append(scenelists()[0])
IMATOP5.append(ITAS5)
IMATOP6.append(scenelists()[0])
IMATOP6.append(ITAS6)
IMATOP7.append(scenelists()[0])
IMATOP7.append(ITAS7)
#this is the list of lists of lists to go through the loop while making breakdown sheets
lolol = []
lolol.append(IMATOP1)
lolol.append(IMATOP2)
lolol.append(IMATOP3)
lolol.append(IMATOP4)
lolol.append(IMATOP5)
lolol.append(IMATOP6)
lolol.append(IMATOP7)

#Function to get list of tagged items from script
#Most common options for bloop are:
#Cast Members, Background Actors, Props, Special Effects, Wardrobe, Music
#Sound, Set Dressing, Special Equipment, Visual Effects, Mechanical Effects, Location
#others can be found in fdx doc 

def taglists(bloop):
	stuff = []
	stuffID = ''
	for dood in root.findall('.TagData/TagCategories/TagCategory'):
		if dood.attrib['Name'] == bloop:
			stuffID = dood.attrib.get('Id')
	for dd in root.findall('.TagData/TagDefinitions/TagDefinition'):
		if dd.attrib['CatId'] == stuffID:
			stuff.append(dd.attrib.get('Label'))
	return sorted(stuff)

#put the list from taglists in a 2D array for importing into GoogleSheets
def numbered_taglists(bloop):
	item_number = []
	bloop_list = []
	for i in range(len(taglists(bloop))):
 		bloop_list.append(taglists(bloop)[i])
 		item_number.append(i+1)
	return list(zip(item_number, bloop_list))

#put this information into a spreadsheet as a stripboard and breakdown sheets
#open the sheet
prepro = client.open("Samantha | Pre Production Package")

#Create a LIST OF ITEMS SHEET for everything 
ad_sheet = prepro.add_worksheet(title="PWS&C", rows = '50', cols = '10')
ad_sheet.update('A1', "Props, Wardrobe, Cast list for {}".format(script))
ad_sheet.update('A3', numbered_taglists("Cast Members"))
ad_sheet.update('C3', numbered_taglists("Props"))
ad_sheet.update('E3', numbered_taglists("Wardrobe"))
ad_sheet.update('G3', numbered_taglists("Special Effects"))

with batch_updater(ad_sheet.spreadsheet) as batch:
	batch.set_column_widths(ad_sheet, [('A', 30), ('B', 125), ('C', 30), ('D', 250), ('E', 30),
		('F', 100), ('G', 30), ('H', 400)])

#Making BREAKDOWN SHEETS for each scene
for i in range(len(sceneinfo())):
#create a breakdown sheet for each scene
	breakdown_sheet = prepro.add_worksheet(title="No. {}".format(sceneinfo()[i][0]), rows = '50', cols = '15')
	breakdown_sheet.update('A1', lolol[i])
 	# format the breakdown sheet
	batch = batch_updater(breakdown_sheet.spreadsheet)
	batch.format_cell_range(breakdown_sheet,'A1:E1', cellFormat(textFormat=textFormat(bold=True)))
	batch.set_row_height(breakdown_sheet, '1', 42)
	batch.set_column_widths(breakdown_sheet, [('A', 30), ('B', 50), ('C', 50), ('D', 300), ('E', 150)])
	batch.execute()

	#an if conditional for the color of the heading is what I would like
	#THIS DIDN'T WORK, IS THERE A WAY?
	if 'C2' == "INT" and 'E2' == "DAY":
		fmt = cellFormat(backgroundColor = color(1, .98, .5))
		format_cell_ranges(breakdown_sheet, [("A1:E2", fmt)])
	elif 'C2' == "EXT" and 'E2' == "NIGHT":
		fmt = cellFormat(backgroundColor = color(.9, .98, .5))
		format_cell_ranges(breakdown_sheet, [("A1:E2", fmt)])
	elif 'C2' == "EXT" and 'E2' == "DAY":
		fmt = cellFormat(backgroundColor = color(1, .98, .9))
		format_cell_ranges(breakdown_sheet, [("A1:E2", fmt)])
