from openpyxl import Workbook

class Saver: # Base Class Only
	def __init__(self, testScripts, outputLoc, outputName):
		self.testScripts =  testScripts
		self.outputLoc =    outputLoc
		self.outputName =   outputName
		self.workbook =     Workbook()
		self.worksheet =    self.workbook.active
	
	def save(self): # Saves the file to the location specified and with the name also given
		if self.outputLoc.endswith("/"):
			savePath = self.outputLoc + self.outputName
		else:
			savePath = self.outputLoc + "/" + self.outputName
		
		self.workbook.save(savePath)
		print("Saved: " + savePath)