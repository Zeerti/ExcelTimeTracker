import win32gui 
import time
import csv
import os
import wx

												#May not need

class excelTimer:
	def __init__(self):
		self.windows = win32gui

		self.start_time = time.time()

		self.activeExcelWindow = {} 													#Store all excel window names with time elapsed as key
		self.excelWindows = list()  	

		self.activeWindow = None
		self.lastActiveWindow = None
		self.currentDirPath = os.path.dirname(os.path.realpath(__file__))

	def resetTimer(self):
		self.start_time = None
		self.start_time = time.time()

	def readInFromCSV(self):
		with open("ExcelFileTimes.csv", 'r') as file:
			reader = csv.reader(file)
			print(reader)
			for line in reader:
				if line:
					temp = line.pop()
					self.activeExcelWindow[temp] = 0
					print("Found Object in list")
					print(temp)
				#self.activeExcelWindow[line] = 0
		


	

	def checkForNewActiveExcelWindow(self): 				#Return True if there is a new active window.
		if self.activeWindow.find("Excel") != -1:
			if self.lastActiveWindow != self.activeWindow:
				return True

	
	def getCurrentPassedTime(self):
		passedTime = time.time() - self.start_time
		return passedTime
	
	def getActiveWindow(self): #Set new active window, returns prior active window
		self.lastActiveWindow = self.activeWindow
		self.activeWindow = self.activeWindow = self.windows.GetWindowText(self.windows.GetForegroundWindow())
		return self.lastActiveWindow

	
	def updateTimeSpentInExcelDocument(self): #Return False if no new document, True if new document
		print("CURRENT TIME THAT HAS PASSED: {}".format(time.time() - self.start_time))

		self.activeExcelWindow[self.activeWindow] += int(self.getCurrentPassedTime()) 
		self.writeToCSV()
		print("Current Time Spent on Document {}".format(self.activeExcelWindow[self.activeWindow]))
	
	def checkIfNewExcelDocFound(self): #FALSE MEANS NO NEW WINDOW FOUND
			if self.activeWindow.find("Excel") != -1:
				print(self.activeExcelWindow)
				if self.activeWindow in self.activeExcelWindow:
					return False
				else:
					self.activeExcelWindow[self.activeWindow] = 0
					print("\n\n\n Adding New Document")
					return True

	def writeToCSV(self):
		with open("ExcelFileTimes.csv", "w+") as file:
			write = csv.DictWriter(file, self.activeExcelWindow.keys())
			write.writeheader()
			write.writerow(self.activeExcelWindow)


#read in from excel file


if __name__ == "__main__":

	app = excelTimer()
	app.readInFromCSV()
	while True:
		time.sleep(1) 												#Pause for a moment before checking again
		app.getActiveWindow()

		if(app.checkForNewActiveExcelWindow()):
			if(app.checkIfNewExcelDocFound() == False):				#Returns True if new doc was added, false if it was already found in dictionary
				app.updateTimeSpentInExcelDocument()				#Adds time to dictionary if excel doc is the same
				app.resetTimer()
			else:
				app.resetTimer()
																	#Sets active window
		if(app.checkForNewActiveExcelWindow() != True):				#if TRUE a NEW EXCEL window is active
			if(app.checkIfNewExcelDocFound() == False):				#Returns True if new doc was added, false if it was already found in dictionary
				app.updateTimeSpentInExcelDocument()				#Adds time to dictionary if excel doc is the same
				app.resetTimer()
			else:
				app.resetTimer()

		




