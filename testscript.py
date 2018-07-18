try:
    from .step import Step
except Exception: #ImportError
    from step import Step
    
class TestScript:
	def __init__(self, name, description, priority, author, status):
		self.name = name
		self.description = description
		self.priority = priority
		self.author = author
		self.status = status
		
		self.steps = []
		
	def addStep(self, number, stepDescription, expectedResult, expectedRunTime=None):
		self.steps.append(Step(number, stepDescription, expectedResult, expectedRunTime=expectedRunTime))