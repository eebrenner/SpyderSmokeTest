#File: spyderSmokeTest.py
#Description: Python script perform a smoke test on Spyder Metering registers
#Date: July 2nd, 2013
#Author: Marc Brenner

##############################
# Revisions                  #
##############################
#1.2 - July 24 2013 - Maintenance Release, minor updates and added comments
#1.1 - July 10 2013 - Created reusable functions to reduce code
#1.0 - July 2  2013 - Initial release

##############################
# Import                     #
##############################

import serial									#Import for serial functions
import minimalmodbus							#Import for reading registers

import xlwt										#Import for excel functions
from xlrd import open_workbook
from xlutils.copy import copy

import time
from time import localtime, strftime

from win32com.client import Dispatch

import os
import sys
import gc										#Garbage collection
import math
import random

##############################
# Serial Setup               #
##############################
port = 10
#port = input('Please enter the USB port number to read from. \n')		#Prompt user for serial port
port = port - 1
address = 1	#input('Please enter the address of the device. \n')		#Prompt user for device address
baud = 115200 #input('Please enter the baud rate of the device. \n')	#Prompt user for baud rate

instrument = minimalmodbus.Instrument(port, address)					#Open device via serial port
instrument.serial.baudrate = baud										#Set Baud Rate
instrument.serial.timeout = .1											#Set timeout

##############################
# Write to Excel             #
##############################

path = os.path.join(os.path.dirname(sys.executable),'spyderSmokeTest.xls')	#Set [path for excel file
originalBook = open_workbook(path)						#Opens file to be modified
newBook = copy(originalBook)							#Creates copy of file
sheet = newBook.get_sheet(0)							#Retrieves an existing sheet to edit

##############################
#Declarations/Initializations#
##############################

gc.disable()											#Turn off garbage collector to speed up list writing

print "+------------------------------------------------------------------------------+"
print "|+----------------------------------------------------------------------------+|"
print "||                         Spyder Smoke Test 1.2                              ||"
print "||----------------------------------------------------------------------------||"
print "||            +---+---+---+---+---+---+---+---+---+---+--------+              ||"
print "||            |   |   |   |   |   |   |   |   |   |   |        |              ||"
print "||            |   |   |   |   |   |   |   |   |   |   |        |              ||"
print "||            |   |   |   |   |   |   |   |   |   |   |        |              ||"
print "||            |   |   |   |   |   |   |   |   |   |   |        |              ||"
print "||            |   |   |   |   |   |   |   |   |   |   |        |              ||"
print "||            |<--+<--+<--+<--+<--+<--+<--+<--+<--+<--|        |              ||"
print "||            | +   +   +   +   +   +   +   +   +   + |        |              ||"
print "||            | |   |   |   |   |   |   |   |   |   | |        |              ||"
print "||            | |   |   |   |   |   |   |   |   |   | |        |              ||"
print "||            | v   v   v   v   v   v   v   v   v   v |        |              ||"
print "||            +---------------------------------------+--------+              ||"
print "|+----------------------------------------------------------------------------+|"
print "+------------------------------------------------------------------------------+"

print '\n'

print "This script will compare values of a single meter from Per Channel Data, "
print "Card Data, Virtual PerParam, Virtual PerTenant, and Virtual Display registers.  "
print "Open the Excel Sheet after running the script to look at values." 
print '\n'
print "Shading:"
print "	Dark green 	: values pass check against theoretical"
print "	Red  		: values do not pass check against theoretical"
print "	Light green 	: values match the other registers"
print "	Yellow 		: values don't match the other registers"
print '\n'

#get input from user
#numberOfMeters = input('\nPlease enter the number of meters in system. \n')		#Use this for random meter, user needs a special hardware config
cardUnderTest = input('\nPlease enter the card number. \n')
channelUnderTest = input('\nPlease enter the channel number. \n')
current = input('\nPlease enter the input current for comparison. \n')
currentAngle = input('\nPlease enter the input current angle for comparison. \n')
voltage = input('\nPlease enter the input voltage for comparison. \n')
voltageAngle = input('\nPlease enter the input voltage angle for comparison. \n')
checkValuesPlease = input('\nWould you like the values checked? (1 = yes) \n')
meterPhases = input('\nPlease enter the number of phases in the meter under test (3/2/1). \n')
print '\n'

#numberOfMeters = 60
numberOfCards = 10
numberOfChannels = 6

#Hardcoded values:
#current = 4
#currentAngle = 175
#voltage = 115
#voltageAngle = 0

#Theoretical Values
currentAngle = math.radians(currentAngle)
voltageAngle = math.radians(voltageAngle)

watts = (current * voltage) * (math.cos(voltageAngle - currentAngle))
vars = (current * voltage) * (math.sin(voltageAngle - currentAngle))
vas = math.sqrt(math.pow(watts,2) + math.pow(vars,2))
pf = math.cos(voltageAngle - currentAngle)

if (voltageAngle == currentAngle):
	pf = pf
elif (voltageAngle < currentAngle):
	pf = -pf
elif (voltageAngle > currentAngle):
	pf = pf

#Set the compare values for individual channel values
if (meterPhases == 3):
	#Limit values
	low = 0.95
	high = 1.05
	compareWatts = 3 * watts
	compareVars = 3 * vars
	compareVas = 3 * vas
elif (meterPhases == 2):
	#Limit values
	low = 0.97
	high = 1.03
	compareWatts = 2 * watts
	compareVars = 2 * vars
	compareVas = 2 * vas
else:
	#Limit values
	low = 0.98
	high = 1.02
	compareWatts = 1 * watts
	compareVars = 1 * vars
	compareVas = 1 * vas

#Other Variables
#currentTimeEpoch = time.time()							#Grab Epoch time
#currentTimeFormatted = time.gmtime()					#Grab Formatted time
#random.seed(currentTimeEpoch)							#Seed to epoch

#Pick random channel/meter code
#meterUnderTest = random.randint(1,numberOfMeters)		#Grab a random meter within range
#channelUnderTest = random.randint(1,numberOfChannels) 	#Grab a random channel within range
#cardUnderTest = random.randint(1,numberOfCards)		#Grab a random card within range

#Debug output
#print '\n'
#print "currentTimeEpoch", currentTimeEpoch, '\n'
#print "currentTimeFormatted", currentTimeFormatted, '\n'
#print "random.seed", random.seed, '\n'

#value lists
theoreticalValues = []
theoreticalValuesCompare = []
virtualPerParamValues = []
virtualPerTenantValues = []
getPerChannelDataValues = []
getCardDataValues = []
getVirtualDisplayValues = []

checkList = []					#Used to standardize all lists

#style lists
theoreticalStyle = []
virtualPerParamStyle = []
virtualPerTenantStyle = []
PerChannelDataStyle = []
CardDataStyle = []
VirtualDisplayStyle = []

readError = 0

##############################
# Functions                  #
##############################

#Gets register value at address
#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string
def getReg(add, function):
	#Modbus reading code.  Will read values at an address and convert to a specific type given by the function number
	readError = 0
	
	while True:
		try:
			if (function == 1):
				reg = instrument.read_float(add, 3, 2)				#Read float
			elif (function == 2):
				reg = instrument.read_register(add, 0)				#Read one register
			elif (function == 3):
				reg = instrument.read_registers(add, 6, 3)			#Read multiple registers
			elif (function == 4):
				reg = instrument.read_long(add, 3, signed=False)	#Read long
			elif (function == 5):
				reg = instrument.read_string(add, 16, 3)			#Read String
				r = ''
				for i in range(0, len(reg), 2):						#Swap chars in string to resolve endian-ness
					r +=  reg[i+1] + reg[i]
				reg = r
			return reg
			break
		except IOError:								#Checks for miscommunication
			time.sleep(.25)
			print add+1,"IOError ", readError
			reg = 65535								#65535 is an invalid number, will be interpreted as error
		except ValueError:							#Checks for value and crc errors on data
			time.sleep(.25)
			print add+1,"ValueError ", readError
			reg = 65535								#65535 is an invalid number, will be interpreted as error
			
		if (readError < 5):							#Retries on error
			readError = readError+1
		else:
			return 65535
			break
			
def getMeterUnderTest():
	#Passes card and channel to 14500 and returns a meter number
	i = 0
	j = 0
	
	add = 14500 									#hardcoded address
	add = add-1
	
	function = 2
	instrument.write_register(add,cardUnderTest)	#Write Card# Out
	reg = getReg(add, function)						
	print add+1, "Card Number:	", reg
	add = add + 1

	instrument.write_register(add,channelUnderTest)	#Write Channel# out
	reg = getReg(add, function)						
	print add+1, "Channel Number:	", reg
	add = add + 1

	reg = getReg(add, function)						
	meterUnderTest = reg
	print add+1, "Meter Number:	", reg
	add = add + 1
	
	print '\n'

	return meterUnderTest
	
def getTheoretical(meterUnderTest):
	#Appends previously calculated values to the theoretical values
	
	#theoreticalValues displays on excel file
	theoreticalValues.append(meterUnderTest)
	theoreticalValues.append(" ")
	theoreticalValues.append(cardUnderTest)
	theoreticalValues.append(channelUnderTest)
	theoreticalValues.append(" ")
	theoreticalValues.append(" ")
	theoreticalValues.append(current)
	theoreticalValues.append(voltage)
	theoreticalValues.append(" ")
	theoreticalValues.append(" ")
	theoreticalValues.append(" ")
	theoreticalValues.append(" ")
	theoreticalValues.append(watts)
	theoreticalValues.append(vars)
	theoreticalValues.append(vas)
	theoreticalValues.append(pf)
	
	#These values reflect the number of phases and used for comparison purposes.  They do not show on the excel file
	theoreticalValuesCompare.append(meterUnderTest)
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(cardUnderTest)
	theoreticalValuesCompare.append(channelUnderTest)
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(current)
	theoreticalValuesCompare.append(voltage)
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(" ")
	theoreticalValuesCompare.append(compareWatts)
	theoreticalValuesCompare.append(compareVars)
	theoreticalValuesCompare.append(compareVas)
	theoreticalValuesCompare.append(pf)

def getVirtualPerParam(meterUnderTest):
	#Grabs Virtual per Parameter register values 
	i = 0
	j = 0
	
	virtualPerParamValues.append(meterUnderTest)	#Append meter # to list
	add = 20000										#Set Address
	add = add +(meterUnderTest * 16) - 16
	add = add-1

	function = 5									#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string
	reg = getReg(add, function)						#Read Name
	virtualPerParamValues.append(reg)
	#print add+1, reg

	add = 20960 									#hardcoded address
	add = add + (meterUnderTest)-1
	add = add-1

	reg = getReg(add, 2)			 				#Read Card Number
	virtualPerParamValues.append(reg)
	#print add+1, "Card Number:	", reg

	add = 21020 									#hardcoded address
	add = add +(meterUnderTest)-1
	add = add-1
	reg = getReg(add, 2) 							#Read Channel Number
	virtualPerParamValues.append(reg)
	#print add+1, "Channel Number:	", reg
	
	while((len(virtualPerParamValues) < 12)):
		virtualPerParamValues.append(" ")			#create empty value for proper spacing
		
	add = 21080 									#hardcoded address
	add = add +(meterUnderTest * 2) - 2
	add = add-1
	
	while (i<36):
		if ((len(virtualPerParamValues) == 16) or (len(virtualPerParamValues) == 25) or (len(virtualPerParamValues) == 34) or (len(virtualPerParamValues) == 43) or (len(virtualPerParamValues) == 52)):
			virtualPerParamValues.append(" ")			#create empty value for proper spacing
		if (i<20):	
			function = 1
			reg = getReg(add, function)				
			#print add+1, reg
			add = add + 120
		elif (i<28):	
			if (i==20):
				add = 23480 							#hardcoded address
				add = add +(meterUnderTest * 6 ) - 6
				add = add-1
			function = 3
			reg = getReg(add, function)				
			reg = str(reg)
			#print add+1, reg
			add = add + 360
		elif (i<38):
			if (i==28):
				add = 26360 							#hardcoded address
				add = add +(meterUnderTest * 2) - 2
				add = add-1
			function = 4
			reg = getReg(add, function)				
			#print add+1, reg
			add = add + 120
		virtualPerParamValues.append(reg)
		
		i = i + 1
	
def getVirtualPerTenant(meterUnderTest):
	#Grabs Virtual per tenant register values 
	i = 0
	j = 0
	
	virtualPerTenantValues.append(meterUnderTest)	#Append meter # to list
	add = 28000										#Set Address
	add = add + ((meterUnderTest-1) * 122)
	add = add-1

	function = 5									#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string
	reg = getReg(add, function)						#Read Name
	virtualPerTenantValues.append(reg)
	#print add+1, reg

	add = add + 16

	reg = getReg(add, 2)			 				#Read Card Number
	virtualPerTenantValues.append(reg)
	#print add+1, "Card Number:	", reg				#Debug

	add = add + 1
	
	reg = getReg(add, 2) 							#Read Channel Number
	virtualPerTenantValues.append(reg)
	#print add+1, "Channel Number:	", reg			#Debug
	
	while((len(virtualPerTenantValues) < 12)):
		virtualPerTenantValues.append(" ")			#create empty value for proper spacing
		
	add = add + 1
	
	while (i<36):
		if ((len(virtualPerTenantValues) == 16) or (len(virtualPerTenantValues) == 25) or (len(virtualPerTenantValues) == 34) or (len(virtualPerTenantValues) == 43) or (len(virtualPerTenantValues) == 52)):
			virtualPerTenantValues.append(" ")		#create empty value for proper spacing
		if (i<20):	
			function = 1
			reg = getReg(add, function)					
			#print add+1, reg
			add = add + 2
		elif (i<28):	
			function = 3
			reg = getReg(add, function)					
			reg = str(reg)
			#print add+1, reg
			add = add + 6
		elif (i<38):
			function = 4
			reg = getReg(add, function)					
			#print add+1, reg
			add = add + 2
		virtualPerTenantValues.append(reg)
		i = i + 1
	
def getPerChannelData():
	#Grabs Per Parameter register values 
	i = 0
	j = 0
	
	add = 14500 									#hardcoded address
	add = add-1
	
	function = 2
	instrument.write_register(add,cardUnderTest)	#Write Card# Out
	reg = getReg(add, function)						
	#print add+1, "Card Number:	", reg				#Debug
	add = add + 1

	instrument.write_register(add,channelUnderTest)	#Write Channel# out
	reg = getReg(add, function)						
	#print add+1, "Channel Number:	", reg			#Debug
	add = add + 1

	reg = getReg(add, function)						
	getPerChannelDataValues.append(reg)
	meterUnderTest = reg
	#print add+1, "Meter Number:	", reg			#Debug
	add = add + 1
	
	getPerChannelDataValues.append(" ")				#create empty value for proper spacing
	
	reg = getReg(14499, function)					#Read back card number
	getPerChannelDataValues.append(reg)
	
	reg = getReg(14500, function)					#Read back channel number
	getPerChannelDataValues.append(reg)
	
	getPerChannelDataValues.append(" ")				#create empty value for proper spacing
	
	while (i<9):
		if ((len(getPerChannelDataValues) == 8) or (len(getPerChannelDataValues) == 11) or (len(getPerChannelDataValues) == 16)):
			getPerChannelDataValues.append(" ")		#create empty value for proper spacing	
		function = 1
		reg = getReg(add, function)					
		#print add+1, reg
		if (i==3):
			reg = getReg(14517-1, 2)				
		elif (i==4):
			reg = getReg(14518-1, function)			
		else:
			add = add + 2
		getPerChannelDataValues.append(reg)
		i = i + 1
	
#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string

def getCardData(meterUnderTest):
	#Grabs Dard Data register values
	i = 0
	j = 1
	
	#Determines the appropriate channel depending on # of phases
	if (meterPhases == 2):			#2 phase meter...
		if (channelUnderTest == 1 or channelUnderTest == 3 or channelUnderTest == 5):
			channelNumber = 1
		elif (channelUnderTest == 2 or channelUnderTest == 4 or channelUnderTest == 6):
			channelNumber = 2
	elif (meterPhases == 3):		#3 phase meter...
		if (channelUnderTest == 1 or channelUnderTest == 4):
			channelNumber = 1
		elif (channelUnderTest == 2 or channelUnderTest == 5):
			channelNumber = 2
		elif (channelUnderTest == 3 or channelUnderTest == 6):
			channelNumber = 3
	else:							#1 phase meter...
		channelNumber = 1
	
	add = 13390 								#hardcoded address
	add = add-1
	
	reg = getReg(add, 2)			 			#Read Card Number
	
	while True:									#Finds virtual meter and correct channel
		reg = getReg(add, 2)					#Grabs meterUnderTest
		#print add+1, meterUnderTest, reg		#Debug
		if (reg == meterUnderTest):				#Compares...
			if (j == channelNumber):			#Checks channel against iteration
				break
			j = j + 1
		add = add + 1
		
	meterOffset = add - 13390 + 1				#Calculates meterOffset for registers
	#print add, meterOffset						#Debug
	
	getCardDataValues.append(reg)
	#print add+1, "Meter Number:	", reg		#Debug
	
	while((len(getCardDataValues) < 5)):
		getCardDataValues.append(" ")			#create empty value for proper spacing
		
	function = 1
	
	add = 13450 + meterOffset * 2 				#meterOffset * 2 for floats
	add = add-1
	
	while (i<9):
		if ((len(getCardDataValues) == 8) or (len(getCardDataValues) == 11) or (len(getCardDataValues) == 16)):
			getCardDataValues.append(" ")		#create empty value for proper spacing	
		if (i==3):
			addOld = add
			add = 14290-1 + meterOffset
			reg = getReg(add, 2)				
			#print add+1, reg
		elif (i==4):
			add = 14350-1 + meterOffset * 2
			reg = getReg(add, function)			
			#print add+1, reg
			add = addOld
		else:
			reg = getReg(add, function)				
			#print add+1, reg
			add = add + 120
		getCardDataValues.append(reg)
		i = i + 1

def getVirtualDisplay(meterUnderTest):
	#Grabs Virtual Display register values
	i = 0
	j = 0
	
	add = 35500 									#hardcoded address
	add = add-1
	
	function = 2
	instrument.write_register(add,meterUnderTest)	#Write Meter# Out
	reg = getReg(add, function)						
	getVirtualDisplayValues.append(reg)
	#print add+1, "Meter Number:	", reg
	add = add + 1

	function = 5									#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string
	reg = getReg(add, function)						#Read Name
	getVirtualDisplayValues.append(reg)
	#print add+1, reg

	add = add + 16

	reg = getReg(add, 2)			 				#Read Card Number
	getVirtualDisplayValues.append(reg)
	#print add+1, "Card Number:	", reg

	add = add + 1
	
	reg = getReg(add, 2) 							#Read Channel Mask
	getVirtualDisplayValues.append(reg)
	#print add+1, "Channel Number:	", reg
	
	while((len(getVirtualDisplayValues) < 12)):
		getVirtualDisplayValues.append(" ")			#create empty value for proper spacing
		
	add = add + 1
	
	while (i<36):
		if ((len(getVirtualDisplayValues) == 16) or (len(getVirtualDisplayValues) == 25) or (len(getVirtualDisplayValues) == 34) or (len(getVirtualDisplayValues) == 43) or (len(getVirtualDisplayValues) == 52)):
			getVirtualDisplayValues.append(" ")			#create empty value for proper spacing
		if (i<20):	
			function = 1
			reg = getReg(add, function)					
			#print add+1, reg
			add = add + 2
		elif (i<28):	
			function = 3
			reg = getReg(add, function)						
			reg = str(reg)
			#print add+1, reg
			add = add + 6
		elif (i<38):
			function = 4
			reg = getReg(add, function)						
			#print add+1, reg
			add = add + 2
		getVirtualDisplayValues.append(reg)
		i = i + 1

def matchValues(): 
	#Compares the values sets with each other; if they all sets are equal, user only has to view one set.
	i = 0
	#Compare Per Channel Data and Card Data
	while (i<16):								
		if ((i == 0) or (i == 5) or (i == 6) or (i == 7) or (i > 11)): 	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
		####Feature to add: Create limit value as an input rather than hardcoded####
			if ((abs(getPerChannelDataValues[i] - getCardDataValues[i]) <= .02 * abs(getPerChannelDataValues[i])) or (getPerChannelDataValues[i] == getCardDataValues[i])):	#Check for near equality
				PerChannelDataStyle[i] = 5
				CardDataStyle[i] = 5
			else:
				PerChannelDataStyle[i] = 2
				CardDataStyle[i] = 2
		elif ((i == 9) or (i == 10)):									#Watch out for integers
			if ((getPerChannelDataValues[i] == getCardDataValues[i])):	#Check for equality
				PerChannelDataStyle[i] = 5
				CardDataStyle[i] = 5
			else:
				PerChannelDataStyle[i] = 2
				CardDataStyle[i] = 2
		else:
			PerChannelDataStyle[i] = 4
			CardDataStyle[i] = 4
		i = i + 1
		
	i = 0
	while (i<60):								#Compare Per Channel Data, Card Data, Display Data
		if (i < 5):								#Match Meter#, name, card and channel
			if (virtualPerParamValues[i] == virtualPerTenantValues[i] == getVirtualDisplayValues[i]):										#Check for equality
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 5
			elif ((virtualPerParamValues[i] == virtualPerTenantValues[i]) and (virtualPerTenantValues[i] != getVirtualDisplayValues[i])):	#What if the display values are wrong?
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 2
			elif ((virtualPerParamValues[i] == getVirtualDisplayValues[i]) and (getVirtualDisplayValues[i] != virtualPerTenantValues)[i]):	#What if the per tenant values are wrong?
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 2
				VirtualDisplayStyle[i] = 5
			elif ((virtualPerTenantValues[i] == getVirtualDisplayValues[i]) and (getVirtualDisplayValues[i] != virtualPerParamValues[i])):	#What if the per param values are wrong?
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 5
			else:
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 2
				VirtualDisplayStyle[i] = 2
		elif ((i == 0) or (i > 11 and i < 16) or (i > 16 and i <25) or (i > 25 and i < 34)): 		#Match float values
			if ((abs(virtualPerParamValues[i] - virtualPerTenantValues[i]) <= .02 * abs(virtualPerParamValues[i])) or (virtualPerParamValues[i] == virtualPerTenantValues[i])):	#Check for near equality
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				if((abs(virtualPerParamValues[i] - getVirtualDisplayValues[i]) <= .02 * abs(virtualPerParamValues[i])) or (virtualPerParamValues[i] == getVirtualDisplayValues[i])):	#Check for near equality
					VirtualDisplayStyle[i] = 5
				else:
					VirtualDisplayStyle[i] = 2
			else:
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 2
		elif ((i > 34 and i < 43) or (i > 43 and i < 52)):						#Check integers
			if ((virtualPerParamValues[i] == virtualPerTenantValues[i])):		#Check for equality
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				if((virtualPerParamValues[i] == getVirtualDisplayValues[i])):	#Check for equality
					VirtualDisplayStyle[i] = 5
				else:
					VirtualDisplayStyle[i] = 2
			else:
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 2
		else:
			virtualPerParamStyle[i] = 4
			virtualPerTenantStyle[i] = 4
			VirtualDisplayStyle[i] = 4
		i = i + 1	
		
def checkValues(value, style, theory):		
	#Compare values with theoretical and assign style
	i = 5
	while (i<16):														#Check Card Data
		if ((i > 5 and i < 9) or (i > 11 and i < 16)):					#Check current, voltage, watt, var, va, and pf
			if (value[i] == " "):
				style[i] = 4
			elif (theory[i] == value[i] ):
				style[i] = 1
			elif (theory[i] < 0):										#For negative values
				if (theory[i] * high <= value[i] <= theory[i] * low):	#Makes sure value is within range
					style[i] = 1
				else:
					style[i] = 3
			elif (theory[i] > 0):										#For postive values
				if (theory[i] * low <= value[i] <= theory[i] * high):	#Makes sure value is within range
					style[i] = 1
				else:
					style[i] = 3	
		i = i + 1	
		
def printValues(value, style, columnNum):
	#Prints lists to excel file
	j = 0
	
	#Set styles
	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	stNull = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
	stGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green;')	#Excel Formatting
	stLightGreen = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')	
	stYellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')	
	stRed = xlwt.easyxf('pattern: pattern solid, fore_colour red;')	
	stWhite = xlwt.easyxf('pattern: pattern no_fill, fore_colour white;')		
	
	#Print Values
	for item in value:													#Print entire list to terminal
		if (style[j] == 0):
			sheet.write(j+1, columnNum, item, stNull)					#Write the register value into cell 
		elif (style[j] == 1):
			sheet.write(j+1, columnNum, item, stGreen)					 
		elif (style[j] == 2):
			sheet.write(j+1, columnNum, item, stYellow)					 
		elif (style[j] == 3):
			sheet.write(j+1, columnNum, item, stRed)					
		elif (style[j] == 5):
			sheet.write(j+1, columnNum, item, stLightGreen)				
		else:
			sheet.write(j+1, columnNum, item, stWhite)					
		#print j+1, item
		j = j + 1
		
def setupExcel():
	#Format the excel spreadsheet for an easy to read appearance
	i = 0
	
	#Set style patterns for excel
	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	stNull = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
	stGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green;')	
	stLightGreen = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')	
	stYellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')	
	stRed = xlwt.easyxf('pattern: pattern solid, fore_colour red;')	
	stWhite = xlwt.easyxf('pattern: pattern solid, fore_colour white;')		
	
	#Set column widths Columns
	while (i<7):
		col = sheet.col(i)							#Grab column i
		col.width = 256 * 20						#Set column width
		i = i + 1
		
	#Set First Row Headers
	sheet.write(0, 1, "Theoretical")		
	sheet.write(0, 2, "Per Channel Data")		
	sheet.write(0, 3, "Card Data")				
	sheet.write(0, 4, "Virtual Per Parameter")	
	sheet.write(0, 5, "Virtual Per Tenant")		
	sheet.write(0, 6, "Virtual Display")	
	
	#Set First Column Labels	
	sheet.write(1, 0, "Meter#")						
	sheet.write(2, 0, "Customer Name")
	sheet.write(3, 0, "Card")
	sheet.write(4, 0, "Channel")					
	sheet.write(5, 0, " ", stNull)							
	sheet.write(6, 0, "Load")						
	sheet.write(7, 0, "Current")					
	sheet.write(8, 0, "Voltage")					
	sheet.write(9, 0, " ", stNull)							
	sheet.write(10, 0, "CT ID")						
	sheet.write(11, 0, "CT Rating")	
	sheet.write(12, 0, " ", stNull)		
	sheet.write(13, 0, "Watts")						
	sheet.write(14, 0, "var")						
	sheet.write(15, 0, "VA")						
	sheet.write(16, 0, "PF")	
	sheet.write(17, 0, " ", stNull)	
	sheet.write(18, 0, "F Demand")					
	sheet.write(19, 0, "R Demand")					
	sheet.write(20, 0, "Q1 var")		
	sheet.write(21, 0, "Q2 var")	
	sheet.write(22, 0, "Q3 var")	
	sheet.write(23, 0, "Q4 var")
	sheet.write(24, 0, "Q1/Q4 VA")	
	sheet.write(25, 0, "Q2/Q3 VA")		
	sheet.write(26, 0, " ", stNull)	
	sheet.write(27, 0, "pk F Demand")					
	sheet.write(28, 0, "pk R Demand")					
	sheet.write(29, 0, "pk Q1 var")		
	sheet.write(30, 0, "pk Q2 var")	
	sheet.write(31, 0, "pk Q3 var")	
	sheet.write(32, 0, "pk Q4 var")
	sheet.write(33, 0, "pk Q1/Q4 VA")	
	sheet.write(34, 0, "pk Q2/Q3 VA")	
	sheet.write(35, 0, " ", stNull)	
	sheet.write(36, 0, "pk F Demand TS")					
	sheet.write(37, 0, "pk R Demand TS")					
	sheet.write(38, 0, "pk Q1 var TS")		
	sheet.write(39, 0, "pk Q2 var TS")	
	sheet.write(40, 0, "pk Q3 var TS")	
	sheet.write(41, 0, "pk Q4 var TS")
	sheet.write(42, 0, "pk Q1/Q4 VA TS")	
	sheet.write(43, 0, "pk Q2/Q3 VA TS")
	sheet.write(44, 0, " ", stNull)	
	sheet.write(45, 0, "pk F Energy")					
	sheet.write(46, 0, "pk R Energy")					
	sheet.write(47, 0, "pk Q1 var Energy")		
	sheet.write(48, 0, "pk Q2 var Energy")	
	sheet.write(49, 0, "pk Q3 var Energy")	
	sheet.write(50, 0, "pk Q4 var Energy")
	sheet.write(51, 0, "pk Q1/Q4 Energy")	
	sheet.write(52, 0, "pk Q2/Q3 Energy")
	
	i = 0
	while (i<8):								#Write grey row spacers 
		sheet.write(5, i, " ", stNull)
		sheet.write(9, i, " ", stNull)			 
		sheet.write(12, i, " ", stNull)			
		sheet.write(17, i, " ", stNull)			
		sheet.write(26, i, " ", stNull)			 
		sheet.write(35, i, " ", stNull)			
		sheet.write(44, i, " ", stNull)			 
		sheet.write(53, i, " ", stNull)			
		i = i + 1

def standardizeLists(checkList, sOption):
	#Standardize lists to a specified length, python does not like it when lists are not the same size
	while (len(checkList) <= 65):
		if (sOption == 0):
			checkList.append(" ")
		elif (sOption == 1):
			checkList.append(4)
	return checkList
	
##############################
# Main                       #
##############################
while True:
	try:
		meterUnderTest = getMeterUnderTest()	#Function to grab meter under test
		getTheoretical(meterUnderTest)			#Function to grab Theoretical values
		getPerChannelData()						#Function to grab Per Channel Data registers
		getCardData(meterUnderTest)
		getVirtualPerParam(meterUnderTest)		#Function to grab Virtual Per Parameter registers
		getVirtualPerTenant(meterUnderTest)		#Function to grab Virtual Per Tenant registers
		getVirtualDisplay(meterUnderTest)		#Function to grab Virtual Display registers
		break
	except IOError:								#Checks for miscommunication on main program
		time.sleep(.25)
		print "Main IOError "
	except ValueError:							#Checks for value and crc errors on main program
		time.sleep(.25)
		print "Main ValueError "

#Standardize lists to avoid length issue
theoreticalValues = standardizeLists(theoreticalValues, 0)
theoreticalValuesCompare = standardizeLists(theoreticalValuesCompare, 0)
theoreticalStyle = standardizeLists(theoreticalStyle, 1)
getPerChannelDataValues = standardizeLists(getPerChannelDataValues, 0)
PerChannelDataStyle = standardizeLists(PerChannelDataStyle, 1)
getCardDataValues = standardizeLists(getCardDataValues, 0)
CardDataStyle = standardizeLists(CardDataStyle, 1)
virtualPerParamValues = standardizeLists(virtualPerParamValues, 0)
virtualPerParamStyle = standardizeLists(virtualPerParamStyle, 1)
virtualPerTenantValues = standardizeLists(virtualPerTenantValues, 0)
virtualPerTenantStyle = standardizeLists(virtualPerTenantStyle, 1)
getVirtualDisplayValues = standardizeLists(getVirtualDisplayValues, 0)
VirtualDisplayStyle = standardizeLists(VirtualDisplayStyle, 1)

#Compares values against each other
matchValues()			
			
#Checks values against theoreticals
if (checkValuesPlease == 1):		
	checkValues(getPerChannelDataValues, PerChannelDataStyle, theoreticalValues)
	checkValues(getCardDataValues, CardDataStyle, theoreticalValues)
	checkValues(virtualPerParamValues, virtualPerParamStyle, theoreticalValuesCompare)
	checkValues(virtualPerTenantValues, virtualPerTenantStyle, theoreticalValuesCompare)
	checkValues(getVirtualDisplayValues, VirtualDisplayStyle, theoreticalValuesCompare)

#Print values to excel with STYLE
printValues(theoreticalValues, theoreticalStyle, 1)
printValues(getPerChannelDataValues, PerChannelDataStyle, 2)
printValues(getCardDataValues, CardDataStyle, 3)
printValues(virtualPerParamValues, virtualPerParamStyle, 4)
printValues(virtualPerTenantValues, virtualPerTenantStyle, 5)
printValues(getVirtualDisplayValues, VirtualDisplayStyle, 6)

#Setup the excel files
#Do this after the printValues for proper apperance
setupExcel()

print "DONE! Open the excel file to check values."

##############################
# Close                      #
##############################

newBook.save('spyderSmokeTest.xls')					#Saves workbook

os.system("start scalc.exe spyderSmokeTest.xls") 	#Open workbook in open office automatically

#Use the following code to open workbook in Excel automatically
#book = Dispatch('Excel.Application')
#book.Visible = 1
#path = os.path.join(os.getcwd(),'spyderSmokeTest.xls')
#book.Workbooks.Open(path)
#book.Worksheets("Smoke Test").Activate()

instrument.serial.close()							#Close Communication