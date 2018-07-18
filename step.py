class Step:
	def __init__(self, number, stepDescription, expectedResult, expectedRunTime=None):
			self.number = "Step " + str(number)
			self.description = stepDescription
			self.expectedResult = expectedResult
			self.expectedRunTime = expectedRunTime