import sys, os, pyodbc, configparser, winsound
from collections import deque
from datetime import datetime

progName = 'Serial Scanner 40000 - Unreleased'
ver = '1.3.0'

#find config file
print("*** Loading config ***")
dirList = os.listdir(os.getcwd())
iniCount = 0
for i in dirList:
	#count ini
	if i.endswith(".ini"):
		iniCount = iniCount + 1

if iniCount == 1:
	for i in dirList:
		if i.endswith(".ini"):
			config = i
elif iniCount < 1:
	print("No config file discovered. Please take care of that, champ.")
	input("Exiting...")
	sys.exit()
else:
	print("More than 1 config file discovered. Please reduce the clutter, mate!")
	input("Exiting...")
	sys.exit()
	
print("*** {0} discovered ***".format(config))

#prepare ini parser
ini = configparser.ConfigParser()
ini.optionxform = str
ini.read(config)

#check version of config
try:
	configVer = ini['nfo']['sftVer'][:3]
except KeyError:
	print("Your ini has no version reference to verify against. So, fail.")
	input("Exiting...")
	sys.exit()
else:
	if configVer == ver[:3]:
		print("*** {0} loaded ***".format(config))
	else:
		print("Config file is not version compatible. Expected {0}, Received {1}.".format(ver[:3],configVer))
		input("Exiting...")
		sys.exit()

#db connect setup
try:
	dbType = ini['nfo']['dbType']
except KeyError:
	print("Your ini has no db type reference. So, fail.")
	input("Exiting...")
	sys.exit()
else:
	if dbType == 'mdb':
		try:
			dir = ini['nfo']['mdbPath']
			dbq = ini['nfo']['DBQ']
			dbTbl = ini['nfo']['table']
			cnxnStr='DRIVER={Driver do Microsoft Access (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};DriverId=25;DefaultDir=' + dir +';DBQ=' + dbq
			cnxn = pyodbc.connect(cnxnStr)
			cur = cnxn.cursor()
		except DatabaseError:
			print("DB connect fail.")
			input("Exiting...")
			sys.exit()
		else:
			print("*** Connected to {0} at {1} ***\n".format(dbTbl,dir))
	elif dbType == 'sql':
		try:
			dbq = ini['nfo']['DBQ']
			dbTbl = ini['nfo']['table']
			server = ini['nfo']['server']
			UID = ini['nfo']['UID']
			pswd = ini['nfo']['pswd']
			cnxnStr='DRIVER={SQL Server};SERVER={0};DATABASE={1};UID={2};PWD={3}'.format(server,dbq,UID,pswd)
			cnxn = pyodbc.connect(cnxnStr)
			cur = cnxn.cursor()
		except DatabaseError:
			print("DB connect fail.")
			input("Exiting...")
			sys.exit()
		else:
			print("*** Connected to {0} in {1} on {2} ***\n".format(dbTbl,dbq,server))
	else:
		print("Db type {0} is unsupported.".format(dbType))
		input("Exiting...")
		sys.exit()

opList = ["skip","quit","done","status"] #to filter op inputs so they arent screened through the DB fetch
mode = 0   #flag for program function modes
timeNow = datetime.strftime(datetime.now(), '%Y%m%d_%H%M%S') #timevar for report titles

#convert ini strings to bool
def str_2_bool(string):
	if string.lower() == 'true':
		return True
	else: 
		return False

#create scan objs to store data
class MakeItem():
	def __init__(self,id,attString):
		self.id = id
		attList = attString.split(' : ')
		self.isSerial = str_2_bool(attList[0])
		self.isMask = str_2_bool(attList[1])
		if len(attList) > 2:
			self.startMask = attList[2]
			self.lengthMask = int(attList[3])
		self.serial = ""
		self.status = ""
		
	def __repr__(self):
		return self.serial
		
	def __len__(self):
		return len(self.serial)
		
	def getScan(self,value,passVal):
		self.serial = value
		self.status = passVal

#func for scan confirmation sound
def conf_sound(type):
	if type == 1: #positive confirmation
		ding = "C:\\Windows\\Media\\Windows Hardware Insert.wav"
		winsound.PlaySound(ding, winsound.SND_FILENAME|winsound.SND_ASYNC)
	elif type == 2: #negative confirmation
		fail = "C:\\Windows\\Media\\Windows Hardware Fail.wav"
		winsound.PlaySound(fail, winsound.SND_FILENAME|winsound.SND_ASYNC)

#check DB for coreSerial (index)
def check_index(value):
	#fetch SN row for validate mode
	selRow = "SELECT [coreSerial] FROM {0};".format(dbTbl)
	cur.execute(selRow)
	dbCoreSn = cur.fetchall()
	for i in dbCoreSn:
		if value in i:
			return 'Success'
						
#check DB for value
def check_db_cross(value,l):
	crossCheckStr = "SELECT {0} FROM {1} WHERE [coreSerial]=(?);".format(l,dbTbl)
	params = (mainDict['coreSerial'].serial)
	cur.execute(crossCheckStr,params)
	cross = cur.fetchone()
	if value in cross:
		return 'Match'
	else:
		return 'Fail'

#define input mask func
def input_mask(listItem,value):
	mStart = mainDict[listItem].startMask
	mLength = mainDict[listItem].lengthMask
	if not value.startswith(mStart):
		print("{0} is not a valid {1}\n{1} should begin with {2}\n".format(value,listItem,mStart))
		conf_sound(2)
		return 'Fail'
	elif len(value) != mLength:
		print("{0} is an invalid {1}\n{2} characters received. Expected {3} characters\n".format(value,listItem,len(value),mLength))
		conf_sound(2)
		return 'Fail'
	else:
		print(listItem, "Accepted:",value,"\n")
		conf_sound(1)
		return 'Success'	
		
#check DB for dupes
def check_db_dupes(value):
	for row in dbRows:
		if value in row:
			return "Dupe"
	else:
		return "Unique"

#check for session dupes
def check_sess_dupes(value):
	for k in mainDict:
		if value == mainDict[k].serial:
			return "Dupe" , k
	else:
		return "Unique"			
		
#define collection func for entry mode
def get_input(listItem):
	while True:
		try:
			value = input("Enter " + listItem + ": ")
			passVal = 'Undetermined'				
		except KeyError:
			print("Key Error?!")
			continue
		else:
			lowValue = value.lower()
			if lowValue in opList:
				if lowValue == "skip": #skip without writing to dict
					if listItem == 'coreSerial':
						print("Can't skip coreSerial. Figure it out champ!")
						continue
					else:
						print("Item Skipped.\n")
						break
				
				elif lowValue == "quit": #quit
					exit()
				
				elif lowValue == "done": #change hook to exit iter without breaking
					print("Okay, finishing up...\n")
					break			
			
				elif lowValue == "status":
					for i in mainList:
						if mainDict[i].serial != "":
							print("{0} : {1}".format(i,mainDict[i].serial))	
					print("\n")			
					continue
			
			if mode == 1:
				#All non-serialized entries here, before we check dupes
				#check for mask, then proceed
				if mainDict[listItem].isSerial == False:
					if mainDict[listItem].isMask == True: 
						if input_mask(listItem,value) == 'Fail':
							continue
						else:
							break
					else:
						break
			
				#Dupe check for current session
				elif check_sess_dupes(value)[0] == "Dupe":
					conf_sound(2)
					print("This serial has already been scanned for this session:")
					print(check_sess_dupes(value)[1],mainDict[check_sess_dupes(value)[1]].serial,"\n")
					
				#Dupe check for DB
				elif check_db_dupes(value) == "Dupe":
					conf_sound(2)
					if listItem == 'coreSerial':
						print("This Core Serial has been entered before.")
						if update_mode() == 'Yes':
							break
						else:
							continue
					else:
						print("This serial is already in the database. Try Again.")
						continue
				
				#scan has been dupe verified, so check for masking, no first, yes second
				elif mainDict[listItem].isMask == False:
					print(listItem, "Accepted:",value,"\n")
					conf_sound(1)
					break			
				
				#mask for valid input
				else:
					if input_mask(listItem,value) == 'Fail':
						continue
					else:
						break
						
			elif mode == 2:
				if listItem == 'coreSerial':
					indexCheck = check_index(value)
					if indexCheck == 'Success':
						index = value
						print("Core Serial Accepted:",value,"\n")
						passVal = 'Passed'
						conf_sound(1)
						break
					else:
						print("Please scan a valid Core Serial Number.")
						conf_sound(2)
						continue
					
				else: 
					valueCheck = check_db_cross(value,listItem)
					if valueCheck == 'Match':
						print("Scanned",listItem,"matches DB.\n")
						passVal = 'Passed'
						conf_sound(1)
						break
					else:
						print("Scanned",listItem,"does not match DB.\n")
						passVal = 'Failed'
						conf_sound(2)
						break
						
	return value, passVal
	
#DB insert function
def insert_db(v):
	insertRow ="INSERT INTO {0} (coreSerial) VALUES ((?));".format(dbTbl)
	params = (v)
	try:
		cur.execute(insertRow,params)
		cur.commit()
		cnxn.commit()
	except pyodbc.IntegrityError:
		print(v,'oops, int err')

#DB update function
def update_db(k,v,snRef):
	print("Updating {0} : {1}".format(k,v))
	updateRow = "UPDATE {0} SET {1}=(?) WHERE coreSerial=(?);".format(dbTbl,k)
	params = (v,snRef)
	att = 1
	thresh = 5
	while True:
		try:
			cur.execute(updateRow,params)
			cur.commit()
			cnxn.commit()
		except pyodbc.IntegrityError:
			print(k,v,snRef,'oops, int err')
		else:	
			updateCheck = check_db_cross(v,k)
			if updateCheck == 'Match':
				print("DB update {0} : {1}. Success on attempt {2}".format(k,v,att))
				return 'Success'
				break
				
			elif att < thresh:
				print("Fail on attempt {0}.".format(att))
				att = att + 1			
				continue
			
			else:
				dbCont = input("DB update failed after {0} attempts. Would you like to attempt again? (y/n)".format(att))
				if dbCont.lower() in ('n','no'):
					print('DB update skipped moving on to next value.')
					return 'Fail'
					break
						
				elif dbCont.lower() in ('y','yes'):
					thresh = thresh + 5
					print("Okay, trying again...")
					continue
				else:
					thresh = thresh + 5
					print("That wasn't a valid response, since you're apparently drunk, i'll just try again...")
					continue
				
def move_to_db_op():
	while True:
		try:
			cont = input("Dict complete, update database? (y/n)").lower()
		except KeyError:
			print("Key Error?!")
			continue
		else:
			if cont in ('n','no'):
				return 'No'
				break
						
			elif cont in ('y','yes'):
				return 'Yes'
				break
			
			else:
				print("It's a simple question...try again.")
				continue

def update_mode():
	while True:
		try:
			cont = input("Would you like to update values for this Core Serial? (y/n)").lower()
		except KeyError:
			print("Key Error?!")
			continue
		else:
			if cont in ('n','no'):
				return 'No'
				break					
			elif cont in ('y','yes'):
				return 'Yes'
				break		
			else:
				print("It's a simple question...try again.")
				continue

				
def exit():
	print("Cleaning up and quitting.")
	cnxn.close()
	sys.exit()
	
#---------------------------------program start----------------------------------------------------------------------------#
#intro
print("Welcome to {0} v{1}".format(progName,ver))
print("'Quit' to exit.\n")

#Ask user for mode selection
while True:
	try:
		print("Which scanner mode would you like? Entry or Validate?")
		value = input("Select mode: ").lower()
	except KeyError:
		print("Key Error?!")
		continue
	else:
		if value in ('e','entry'):
			mode = 1
			print("Entry Mode Selected\n")
			break
		elif value in ('v','valid','validate'):
			mode = 2
			print("Validate Selected\n")
			break
		elif value == 'quit': #quit
			exit()
		else:
			print("C'mon, just pick a mode, it's not that complicated. \n")
			continue

#generate mainList and mainDict	
mainList = deque([])
mainDict = {}
for k in ini['dbCol']:
		mainList.append(k)
		mainDict.update({k:MakeItem(k,ini['dbCol'][k])})
		
 #Entry mode setup		
if mode == 1:
	#fetch db serials for entry mode dupe check
	selAllStr = "SELECT * FROM {0};".format(dbTbl)
	cur.execute(selAllStr)
	dbRows = cur.fetchall()
	
print("'Status' to display current collection.")
print("'Skip' to pass any item.")
print("'Done' to complete collection early.\n")

while True:
	#iterate through list and grab inputs for dict
	for i in mainList:
		dictValue, passValue = get_input(i)
	
		if dictValue == 'skip' or dictValue == 'done':
			pass
		else:
			mainDict[i].getScan(dictValue,passValue)
		
		if dictValue == 'done':
			break
			
			
	#summarize dict and generate report file
	if int(ini['nfo']['logging']) == 1:
		print("*** Generating Log file ***")
		srlTitle = mainDict['coreSerial'].serial[-9:]
		maxCol = max(len(k) for k in mainDict)
		maxVal = max(len(v) for v in mainDict.values())
		
		if mode == 2:
			fileTitle = srlTitle + '_Validation_' + timeNow + '.txt'
			fileObj = open(fileTitle, 'w')
			fileObj.write('Validated System Serials\n')
			fileObj.write('---------------------------------------------------------\n')
			for i in mainList:
				if mainDict[i].serial != '':
					line = (i.ljust(maxCol+1) +": " + mainDict[i].serial.ljust(maxVal+1) + mainDict[i].status + "\n")
					fileObj.write(line)
				else:
					line = (i.ljust(maxCol+1) +": " + mainDict[i].serial.ljust(maxVal+1) + "\n")
					fileObj.write(line)
			fileObj.close()
		
		elif mode == 1:
			fileTitle = srlTitle + '_Serials_' + timeNow + '.txt'
			fileObj = open(fileTitle, 'w')
			fileObj.write('Scanned Serial Inputs\n')
			fileObj.write('---------------------------------------------------------\n')
			for i in mainList:
				line = (i.ljust(maxCol+1) +": " + mainDict[i].serial.ljust(maxVal+1) + "\n")
				fileObj.write(line)
			fileObj.close()
										
	else:
		print("Summary:")
		for i in mainList:
			print("{0} : {1}".format(i,mainDict[i]))	
	
	#Continue to DB inputs?		
	if mode == 1:
		if move_to_db_op() == "Yes":
		#iterate through dict for db. keys equal columns and values are values. If core serial, do first input, then update that row (redundant, yes, but easier than typing it out)
			for i in mainList:
				if i == "coreSerial":
					if check_db_dupes(value) == "Dupe":
						pass
					else:
						insert_db(mainDict[i].serial)
				elif mainDict[i].serial != "":
					update_db(i,mainDict[i].serial,mainDict['coreSerial'].serial)
				
			print("DB update complete.")		
	
	#Continue main loop for another system?
	while True:
		try:
			mainCont = input("Would you like to scan another unit? (y/n)").lower()
		except KeyError:
			print("Key Error?!")
			continue
		else:
			if mainCont in ('n','no'):
				exit()
			elif mainCont in ('y','yes'):
				print("Next unit...\n")
				break
			else:
				print("Invalid response, bro. Try again.")
				continue

		
cnxn.close()
print("Exiting...")