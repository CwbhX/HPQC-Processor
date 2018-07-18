import yaml
try:
    from .saver import Saver
except Exception: #ImportError
    from saver import Saver

class HPQCSaver(Saver):
	def __init__(self, testScripts, outputLoc, outputName, headerPath):
		super().__init__(testScripts, outputLoc, outputName)
		self.nextStepRow = 0 # Keep track of where we will start adding the next step as to not to overwrite data
		
		self.loadHeader(headerPath)
		self.setupHeader()
	
	
	def loadHeader(self, headerPath): # Load the header for the document
		self.header = yaml.load(open(headerPath, 'r'))
	
	
	def setupHeader(self): # Add the header to the docuement
		for headerElement in self.header:
			title, location = list(headerElement.items())[0]
			self.worksheet.cell(row=location[0], column=location[1]).value = title
			
		self.nextStepRow = 2 # First step will be inserted on row 2
		print("Header set up!")
	
	
	def addTestSteps(self, testScript):
		stepCount = 0
		for step in testScript.steps: # Adds the steps, step per step, following code adds data along the selected row
			self.worksheet.cell(row=self.nextStepRow + stepCount, column=3).value = step.number          # Add step number
			self.worksheet.cell(row=self.nextStepRow + stepCount, column=4).value = step.description     # Add step description
			self.worksheet.cell(row=self.nextStepRow + stepCount, column=5).value = step.expectedResult  # Add expected result for step
			self.worksheet.cell(row=self.nextStepRow + stepCount, column=2).value = step.expectedRunTime # Add expected runtime for step
			stepCount += 1 # Move on to the next row to add the next step
		
		
		#print("Len of steps in: " + str(testScript.name) + " is:  " + str(len(testScript.steps)) + "  Step count: " + str(stepCount) + "  Step Row: " + str(self.nextStepRow))
		self.nextStepRow += stepCount # Reset the nextStepRow once done adding the steps for this testScript so we know where to place the next test
	
	
	def addTestHeader(self, testScript):
		#print("Next Step Row: " + str(self.nextStepRow))
		if testScript.name != None:
			self.worksheet.cell(row=self.nextStepRow, column=1).value =  testScript.name                 # Add test name if exists
		else:
			self.worksheet.cell(row=self.nextStepRow, column=1).value = "Untitled"                       # Add test name to untitled if doesn't exist
		self.worksheet.cell(row=self.nextStepRow, column=18).value = testScript.description              # Add test description
		self.worksheet.cell(row=self.nextStepRow, column=6).value =  testScript.priority                 # Add test priority
		self.worksheet.cell(row=self.nextStepRow, column=7).value =  testScript.author                   # Add test designer
		self.worksheet.cell(row=self.nextStepRow, column=17).value = testScript.status                   # Add test status
		
	
	def addTests(self):
		for testScript in self.testScripts:
			self.addTestHeader(testScript)
			self.addTestSteps(testScript)
		
		print("Wrote tests to sheet!")