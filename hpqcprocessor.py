import yaml, sys, time
try:
    from .loader import Loader
except Exception: #ImportError
    from loader import Loader
    
try:
    from .hpqcsaver import HPQCSaver
except Exception: #ImportError
    from hpqcsaver import HPQCSaver


class HPQCProcessor:
	def __init__(self, directory, configLoc, outputLoc):
		self.directory = directory
		self.outputLoc = outputLoc
	
		self.testScripts = []
		
		self.loadConfig(configLoc)
	
	
	def loadConfig(self, configLoc):
		self.config = yaml.load(open(configLoc, 'r'))
		
		
	def loadData(self):
		loader = Loader(self.directory, self.config)
		loader.load() # Load the scripts from the given directory given the configuration file
		
		self.testScripts = loader.testScripts


	def saveData(self, outputLoc, outputName, outputFormat, saveConfig):
		if outputFormat == "HPQC": # Allows for programming different formats in the future
			self.saveDataToHPQC(outputLoc, outputName, saveConfig) # Make sure this file is in the same directory as run environment


	def saveDataToHPQC(self, outputLoc, outputName, headerPath):
		saver = HPQCSaver(self.testScripts, outputLoc, outputName, headerPath) # Load saver instance
		saver.addTests() # Add tests to the active save sheet
		saver.save()     # Save the sheet and workbook to a file

# python3 hpqcprocessor.py ~/testscripts/ ~/workspace/hpqc_projects/HPQCProcessor/config.yaml ~/workspace/hpqc_projects/HPQCProcessor/hpqcheader.yaml  ~/ hpqctestoutput.xlsx HPQC
if __name__ == "__main__":
	arguments = sys.argv
	spreadsheetsDirectory = arguments[1]
	configFile = arguments[2]
	saveConfig = arguments[3]
	outputDirectory = arguments[4]
	outputName = arguments[5]
	outputFormat = arguments[6]
	
	
	print("\n")
	t = time.time()
	hpqcProcessor = HPQCProcessor(spreadsheetsDirectory, configFile, outputDirectory)
	hpqcProcessor.loadData()
	print("Saving...")
	hpqcProcessor.saveData(outputDirectory, outputName, outputFormat, saveConfig)
	t = time.time() - t
	print("Done, took: %.2fs\n" % t)
	