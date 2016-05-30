#imports

from openpyxl import * 

#root through each user's day 1 and day 2 data

def fileName(letter, ID, day):
	"""Arguments: letter: E or C (experiment/stutter or control)
	ID: participant ID 
	day: 1 or 2

	Returns: string filename of stats

	"""


	#studyGroups = ["C", "E"]
	#controlID = [1, 3, 4, 6, 12, 13, 16, 17]
	#stutterID = [2, 4, 5, 6, 7, 9, 10, 11]
	
	#for group in studyGroups:
	#	if group == "C":
	#		for ID in controlID: 
	#			filename1 = "P" + group + str(ID) + "D1_day_0_ACC.xlsx"
	#			filename2 = "P" + group + str(ID) + "D2_day_1_ACC.xlsx"
	#			print filename1
	#			print filename2
	#	else: 
	#		for ID in stutterID: 
	#			filename1 = "P" + group + str(ID) + "D1_day_0_ACC.xlsx"
	#			filename2 = "P" + group + str(ID) + "D2_day_1_ACC.xlsx"
	#			print filename1
	#			print filename2

	filename = "P" + letter + str(ID) + "D" + str(day) + "_day_" + str((day-1)) + "_ACC" + ".xlsx"
	return filename 

def calculateMeanAccuracy(worksheet, start):
	i = start
	j = start
	meanAccuracies = []
	a = 1
	while a < 7:
		accuracyTempList = []
		while i < j+10:
			coordinate2 = "F" + str(i)
			value =  worksheet.cell(coordinate2).value
			accuracyTempList.append(int(value))
			i = i+1
		num = float(sum(accuracyTempList))
		den = float(len(accuracyTempList))
		mean = num/den
		meanAccuracies.append(mean)
		i = j + 10
		j = j+ 10 
		a = a+1
	return meanAccuracies

#getStats('C', 1,1)

def main():
	filename = fileName('C', 1, 1)
	workbook = load_workbook(filename) 
	worksheet = workbook.active
	#print worksheet.cell("A200").value

	#Note: with current formatting, in column A, cells have value of None if 
	#it's from aforementioned sequence.self
	#OR None if there is just no value there because the column is over

	accuracyList = []
	sequenceList = []
	blockList = [1,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3]
	for i in range(2,184):
		coordinate = "A" + str(i)
		if worksheet.cell(coordinate).value != None:
			sequenceList.append(worksheet.cell(coordinate).value)

	accuracyList.append(calculateMeanAccuracy(worksheet, 2))
	accuracyList.append(calculateMeanAccuracy(worksheet, 63))
	accuracyList.append(calculateMeanAccuracy(worksheet, 124))

main() 
