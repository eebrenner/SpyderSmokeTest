#File: spyderSmokeTest.py
#Description: Python script perform a smoke test on Spyder Metering registers
#Date: July 2nd, 2013
#Author: Marc Brenner

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
import gc
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
instrument.serial.baudrate = baud
instrument.serial.timeout = .1

##############################
# Write to Excel             #
##############################

path = os.path.join(os.path.dirname(sys.executable),'spyderSmokeTest.xls')
originalBook = open_workbook(path)						#Opens file to be modified
newBook = copy(originalBook)							#Creates copy of file
sheet = newBook.get_sheet(0)							#Retrieves an existing sheet to edit

##############################
#Declarations/Initializations#
##############################

gc.disable()											#Turn off garbage collector to speed up list writing

print "+------------------------------------------------------------------------------+"
print "|+----------------------------------------------------------------------------+|"
print "||                         Spyder Smoke Test 1.0                              ||"
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
#timeLimit = input('How many times would you like to run the program? \n')

#numberOfMeters = input('Please enter the number of meters in system. \n')
#cardUnderTest = input('Please enter the card number. \n')
print '\n'
#channelUnderTest = input('Please enter the channel number. \n')
print '\n'
#current = input('Please enter the input current for comparison. \n')
print '\n'
#currentAngle = input('Please enter the input current angle for comparison. \n')
print '\n'
#voltage = input('Please enter the input voltage for comparison. \n')
print '\n'
#voltageAngle = input('Please enter the input voltage angle for comparison. \n')
print '\n'
checkValuesPlease = input('Would you like the values checked? (1 = yes) \n')
print '\n'
meterPhases = input('Please enter the number of phases in the meter under test (3/2/1). \n')
print '\n'

#numberOfMeters = 60
numberOfCards = 10
numberOfChannels = 6

current = 4
currentAngle = 5
voltage = 115
voltageAngle = 0

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
currentTimeEpoch = time.time()						#Grab Epoch time
#currentTimeFormatted = time.gmtime()				#Grab Formatted time
random.seed(currentTimeEpoch)						#Seed to epoch

#meterUnderTest = 1
#meterUnderTest = random.randint(1,numberOfMeters)	#Grab a random meter within range
channelUnderTest = random.randint(1,numberOfChannels) #Grab a random channel within range
cardUnderTest = random.randint(1,numberOfCards)	#Grab a random card within range

#print '\n'
#print "currentTimeEpoch", currentTimeEpoch, '\n'
#print "currentTimeFormatted", currentTimeFormatted, '\n'
#print "random.seed", random.seed, '\n'

#lists
theoreticalValues = []
theoreticalValuesCompare = []
virtualPerParamValues = []
virtualPerTenantValues = []
getPerChannelDataValues = []
getCardDataValues = []
getVirtualDisplayValues = []

checkList = []

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
				for i in range(0, len(reg), 2):						#Swap chars in string
					r += reg[i+1] + reg[i]
				reg = r
			return reg
			break
		except IOError:									#Checks for miscommunication
			time.sleep(.25)
			print add+1,"IOError ", readError
			reg = 65535								#65535 is an invalid number, will be interpreted as error
		except ValueError:
			time.sleep(.25)
			print add+1,"ValueError ", readError
			reg = 65535								#65535 is an invalid number, will be interpreted as error
			
		if (readError < 5):							#Retries on error
			readError = readError+1
		else:
			return 65535
			break
			
def getMeterUnderTest():
	i = 0
	j = 0
	
	add = 14500 										#hardcoded address
	add = add-1
	
	function = 2
	instrument.write_register(add,cardUnderTest)	#Write Card# Out
	reg = getReg(add, function)						#Read floats
	print add+1, "Card Number:	", reg
	add = add + 1

	instrument.write_register(add,channelUnderTest)	#Write Channel# out
	reg = getReg(add, function)						#Read floats
	print add+1, "Channel Number:	", reg
	add = add + 1

	reg = getReg(add, function)						#Read floats
	meterUnderTest = reg
	print add+1, "Meter Number:	", reg
	add = add + 1
	
	print '\n'

	return meterUnderTest
	
def getTheoretical(meterUnderTest):
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
	
	#These values reflect the number of phases
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
			virtualPerParamValues.append(" ")		#create empty value for proper spacing
		if (i<20):	
			function = 1
			reg = getReg(add, function)				#Read floats
			#print add+1, reg
			add = add + 120
		elif (i<28):	
			if (i==20):
				add = 23480 										#hardcoded address
				add = add +(meterUnderTest * 6 ) - 6
				add = add-1
			function = 3
			reg = getReg(add, function)				#Read floats
			reg = str(reg)
			#print add+1, reg
			add = add + 360
		elif (i<38):
			if (i==28):
				add = 26360 										#hardcoded address
				add = add +(meterUnderTest * 2) - 2
				add = add-1
			function = 4
			reg = getReg(add, function)				#Read floats
			#print add+1, reg
			add = add + 120
		virtualPerParamValues.append(reg)
		
		i = i + 1
	
def getVirtualPerTenant(meterUnderTest):
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
	#print add+1, "Card Number:	", reg

	add = add + 1
	
	reg = getReg(add, 2) 							#Read Channel Number
	virtualPerTenantValues.append(reg)
	#print add+1, "Channel Number:	", reg
	
	while((len(virtualPerTenantValues) < 12)):
		virtualPerTenantValues.append(" ")			#create empty value for proper spacing
		
	add = add + 1
	
	while (i<36):
		if ((len(virtualPerTenantValues) == 16) or (len(virtualPerTenantValues) == 25) or (len(virtualPerTenantValues) == 34) or (len(virtualPerTenantValues) == 43) or (len(virtualPerTenantValues) == 52)):
			virtualPerTenantValues.append(" ")			#create empty value for proper spacing
		if (i<20):	
			function = 1
			reg = getReg(add, function)					#Read floats
			#print add+1, reg
			add = add + 2
		elif (i<28):	
			function = 3
			reg = getReg(add, function)					#Read floats
			reg = str(reg)
			#print add+1, reg
			add = add + 6
		elif (i<38):
			function = 4
			reg = getReg(add, function)					#Read floats
			#print add+1, reg
			add = add + 2
		virtualPerTenantValues.append(reg)
		i = i + 1
	
def getPerChannelData():
	i = 0
	j = 0
	
	add = 14500 										#hardcoded address
	add = add-1
	
	function = 2
	instrument.write_register(add,cardUnderTest)	#Write Card# Out
	reg = getReg(add, function)						#Read floats
	#print add+1, "Card Number:	", reg
	add = add + 1

	instrument.write_register(add,channelUnderTest)	#Write Channel# out
	reg = getReg(add, function)						#Read floats
	#print add+1, "Channel Number:	", reg
	add = add + 1

	reg = getReg(add, function)						#Read floats
	getPerChannelDataValues.append(reg)
	meterUnderTest = reg
	#print add+1, "Meter Number:	", reg
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
		reg = getReg(add, function)					#Read floats
		#print add+1, reg
		if (i==3):
			reg = getReg(14517-1, 2)				#Read floats
		elif (i==4):
			reg = getReg(14518-1, function)			#Read floats
		else:
			add = add + 2
		getPerChannelDataValues.append(reg)
		i = i + 1
	
#1 = float, 2 = reg, 3 = regs, 4 = long, 5 = string

def getCardData(meterUnderTest):
	i = 0
	j = 1
	
	if (meterPhases == 2):
		if (channelUnderTest == 1 or channelUnderTest == 3 or channelUnderTest == 5):
			channelNumber = 1
		elif (channelUnderTest == 2 or channelUnderTest == 4 or channelUnderTest == 6):
			channelNumber = 2
	elif (meterPhases == 3):
		if (channelUnderTest == 1 or channelUnderTest == 4):
			channelNumber = 1
		elif (channelUnderTest == 2 or channelUnderTest == 5):
			channelNumber = 2
		elif (channelUnderTest == 3 or channelUnderTest == 6):
			channelNumber = 3
	else:
		channelNumber = 1
	
	#getCardDataValues.append(meterUnderTest)		#Append meter # to list
	add = 13390 									#hardcoded address
	add = add-1
	
	reg = getReg(add, 2)			 				#Read Card Number
	
	while True:
		reg = getReg(add, 2)
		#print add+1, meterUnderTest, reg
		if (reg == meterUnderTest):
			if (j == channelNumber):
				break
			j = j + 1
		add = add + 1
		
	meterOffset = add - 13390 + 1
	#print add, meterOffset	
	
	getCardDataValues.append(reg)
	#print add+1, "Meter Number:	", reg
	
	while((len(getCardDataValues) < 5)):
		getCardDataValues.append(" ")				#create empty value for proper spacing
		
	function = 1
	
	add = 13450 + meterOffset * 2 
	add = add-1
	
	while (i<9):
		if ((len(getCardDataValues) == 8) or (len(getCardDataValues) == 11) or (len(getCardDataValues) == 16)):
			getCardDataValues.append(" ")		#create empty value for proper spacing	
		if (i==3):
			addOld = add
			add = 14290-1 + meterOffset
			reg = getReg(add, 2)				#Read floats
			#print add+1, reg
		elif (i==4):
			add = 14350-1 + meterOffset * 2
			reg = getReg(add, function)		#Read floats
			#print add+1, reg
			add = addOld
		else:
			reg = getReg(add, function)			#Read floats
			#print add+1, reg
			add = add + 120
		getCardDataValues.append(reg)
		i = i + 1

def getVirtualDisplay(meterUnderTest):
	i = 0
	j = 0
	
	add = 35500 										#hardcoded address
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
			reg = getReg(add, function)					#Read floats
			#print add+1, reg
			add = add + 2
		elif (i<28):	
			function = 3
			reg = getReg(add, function)					#Read floats
			reg = str(reg)
			#print add+1, reg
			add = add + 6
		elif (i<38):
			function = 4
			reg = getReg(add, function)					#Read floats
			#print add+1, reg
			add = add + 2
		getVirtualDisplayValues.append(reg)
		i = i + 1

def matchValues(): 				#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	i = 0
	while (i<300):							#Set up style lists
		#theoreticalStyle.append(4)			#Set up style lists
		virtualPerParamStyle.append(0)
		virtualPerTenantStyle.append(0)
		PerChannelDataStyle.append(0)
		CardDataStyle.append(0)
		VirtualDisplayStyle.append(0)
		i = i + 1
		
	#Match values and assign style
	i = 0
	while (i<16):								#Compare Per Channel Data and Card Data
		if ((i == 0) or (i == 5) or (i == 6) or (i == 7) or (i > 11)): 
			if ((abs(getPerChannelDataValues[i] - getCardDataValues[i]) <= .02 * abs(getPerChannelDataValues[i])) or (getPerChannelDataValues[i] == getCardDataValues[i])):	#Check for near equality
				PerChannelDataStyle[i] = 5
				CardDataStyle[i] = 5
			else:
				PerChannelDataStyle[i] = 2
				CardDataStyle[i] = 2
		elif ((i == 9) or (i == 10)):									#Watch out for integers
			if ((getPerChannelDataValues[i] == getCardDataValues[i])):	#Check for near equality
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
		if (i < 5):
			if (virtualPerParamValues[i] == virtualPerTenantValues[i] == getVirtualDisplayValues[i]):	#Check for near equality
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 5
			elif ((virtualPerParamValues[i] == virtualPerTenantValues[i]) and (virtualPerTenantValues[i] != getVirtualDisplayValues[i])):
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 2
			elif ((virtualPerParamValues[i] == getVirtualDisplayValues[i]) and (getVirtualDisplayValues[i] != virtualPerTenantValues)[i]):
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 2
				VirtualDisplayStyle[i] = 5
			elif ((virtualPerTenantValues[i] == getVirtualDisplayValues[i]) and (getVirtualDisplayValues[i] != virtualPerParamValues[i])):
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 5
				VirtualDisplayStyle[i] = 5
			else:
				virtualPerParamStyle[i] = 2
				virtualPerTenantStyle[i] = 2
				VirtualDisplayStyle[i] = 2
		elif ((i == 0) or (i > 11 and i < 16) or (i > 16 and i <25) or (i > 25 and i < 34)): 
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
		elif ((i > 34 and i < 43) or (i > 43 and i < 52)):
			if ((virtualPerParamValues[i] == virtualPerTenantValues[i])):		#Check for near equality
				virtualPerParamStyle[i] = 5
				virtualPerTenantStyle[i] = 5
				if((virtualPerParamValues[i] == getVirtualDisplayValues[i])):	#Check for near equality
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
		
def checkValues():		
	#Compare values and assign style
	i = 5
	while (i<16):								#Check per channel Data
		if ((i > 5 and i < 9) or (i > 11 and i < 16)):					#Check load, current, voltage, watt, var, va, and pf
			if (getPerChannelDataValues[i] == " "):
				PerChannelDataStyle[i] = 4
			elif (theoreticalValues[i] == getPerChannelDataValues[i] ):
				PerChannelDataStyle[i] = 1
			elif (theoreticalValues[i] < 0):
				if (theoreticalValues[i] * high <= getPerChannelDataValues[i] <= theoreticalValues[i] * low):
					PerChannelDataStyle[i] = 1
				else:
					PerChannelDataStyle[i] = 3
			elif (theoreticalValues[i] > 0):
				if (theoreticalValues[i] * low <= getPerChannelDataValues[i] <= theoreticalValues[i] * high):
					PerChannelDataStyle[i] = 1
				else:
					PerChannelDataStyle[i] = 3	
		i = i + 1	
		
	i = 5
	while (i<16):								#Check Card Data
		if ((i > 5 and i < 9) or (i > 11 and i < 16)):					#Check load, current, voltage, watt, var, va, and pf
			if (getCardDataValues[i] == " "):
				CardDataStyle[i] = 4
			elif (theoreticalValues[i] == getCardDataValues[i] ):
				CardDataStyle[i] = 1
			elif (theoreticalValues[i] < 0):
				if (theoreticalValues[i] * high <= getCardDataValues[i] <= theoreticalValues[i] * low):
					CardDataStyle[i] = 1
				else:
					CardDataStyle[i] = 3
			elif (theoreticalValues[i] > 0):
				if (theoreticalValues[i] * low <= getCardDataValues[i] <= theoreticalValues[i] * high):
					CardDataStyle[i] = 1
				else:
					CardDataStyle[i] = 3	
		i = i + 1	

	i = 5
	while (i<16):								#Check Per Param Data
		if ((theoreticalValuesCompare[i] == " ") or (virtualPerParamValues[i] == " ")):
			virtualPerParamStyle[i] = 4
		elif (theoreticalValuesCompare[i] < 0):
			if (theoreticalValuesCompare[i] * high <= virtualPerParamValues[i] <= theoreticalValuesCompare[i] * low):
				virtualPerParamStyle[i] = 1
			else:
				virtualPerParamStyle[i] = 3
		elif (theoreticalValuesCompare[i] > 0):
			if (theoreticalValuesCompare[i] * low <= virtualPerParamValues[i] <= theoreticalValuesCompare[i] * high):
				virtualPerParamStyle[i] = 1
			else:
				virtualPerParamStyle[i] = 3	
		i = i + 1	
		
	i = 5
	while (i<16):								#Check Virtual per Tenant Data
		if ((theoreticalValuesCompare[i] == " ") or (virtualPerTenantValues[i] == " ")):
			virtualPerTenantStyle[i] = 4
		elif (theoreticalValuesCompare[i] < 0):
			if (theoreticalValuesCompare[i] * high <= virtualPerTenantValues[i] <= theoreticalValuesCompare[i] * low):
				virtualPerTenantStyle[i] = 1
			else:
				virtualPerTenantStyle[i] = 3
		elif (theoreticalValuesCompare[i] > 0):
			if (theoreticalValuesCompare[i] * low <= virtualPerTenantValues[i] <= theoreticalValuesCompare[i] * high):
				virtualPerTenantStyle[i] = 1
			else:
				virtualPerTenantStyle[i] = 3	
		i = i + 1	
	
	i = 5
	while (i<16):								#Check Virtual Display Data
		if ((theoreticalValuesCompare[i] == " ") or (getVirtualDisplayValues[i] == " ")):
			VirtualDisplayStyle[i] = 4
		elif (theoreticalValuesCompare[i] < 0):
			if (theoreticalValuesCompare[i] * high <= getVirtualDisplayValues[i] <= theoreticalValuesCompare[i] * low):
				VirtualDisplayStyle[i] = 1
			else:
				VirtualDisplayStyle[i] = 3
		elif (theoreticalValuesCompare[i] > 0):
			if (theoreticalValuesCompare[i] * low <= getVirtualDisplayValues[i] <= theoreticalValuesCompare[i] * high):
				VirtualDisplayStyle[i] = 1
			else:
				VirtualDisplayStyle[i] = 3	
		i = i + 1
		
def printValues(value, style, columnNum):
	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	stNull = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
	stGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green;')	#Excel Formatting
	stLightGreen = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')	
	stYellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')	
	stRed = xlwt.easyxf('pattern: pattern solid, fore_colour red;')	
	stWhite = xlwt.easyxf('pattern: pattern no_fill, fore_colour white;')		
	
	#Print Values
	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	j = 0
	
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
	#0 = gray, 1 = green, 2 = yellow, 3 = red, 4 = white
	stNull = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
	stGreen = xlwt.easyxf('pattern: pattern solid, fore_colour green;')	#Excel Formatting
	stLightGreen = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')	
	stYellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')	
	stRed = xlwt.easyxf('pattern: pattern solid, fore_colour red;')	
	stWhite = xlwt.easyxf('pattern: pattern solid, fore_colour white;')		
	
	#Set column widths Columns
	i = 0
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
	while (i<8):
		sheet.write(5, i, " ", stNull)			#Write header 
		sheet.write(9, i, " ", stNull)			#Write header 
		sheet.write(12, i, " ", stNull)			#Write header 
		sheet.write(17, i, " ", stNull)			#Write header 
		sheet.write(26, i, " ", stNull)			#Write header 
		sheet.write(35, i, " ", stNull)			#Write header 
		sheet.write(44, i, " ", stNull)			#Write header 
		sheet.write(53, i, " ", stNull)			#Write header 
		i = i + 1

def standardizeLists(checkList, sOption):
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
	except IOError:									#Checks for miscommunication
		time.sleep(.25)
		print "Main IOError "
	except ValueError:									#Checks for miscommunication
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

matchValues()
if (checkValuesPlease == 1):
	checkValues()

printValues(theoreticalValues, theoreticalStyle, 1)
printValues(getPerChannelDataValues, PerChannelDataStyle, 2)
printValues(getCardDataValues, CardDataStyle, 3)
printValues(virtualPerParamValues, virtualPerParamStyle, 4)
printValues(virtualPerTenantValues, virtualPerTenantStyle, 5)
printValues(getVirtualDisplayValues, VirtualDisplayStyle, 6)

setupExcel()

print "DONE! Open the excel file to check values."
##############################
# Close                      #
##############################

newBook.save('spyderSmokeTest.xls')				#Saves workbook

#book = Dispatch('Excel.Application')
#book.Visible = 1
#path = os.path.join(os.getcwd(),'spyderSmokeTest.xls')
#book.Workbooks.Open(path)
#book.Worksheets("Smoke Test").Activate()
os.system("start scalc.exe spyderSmokeTest.xls") #Open workbook

instrument.serial.close()