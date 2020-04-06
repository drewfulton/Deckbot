#----------------------------------------------------------------------------
# Name: views.py
# Purpose: 	Provide User Interface
# Author:	Drew Fulton
# Created:	April 4, 2020


class DeckbotCLI(object):
	''' Creates a simple CLI option for creating user input to select a company
	'''
	def ___init___(self):
		''' Display list of company options and ask for Input from User in CLI
		'''
		self.company_list = company_list
		
	def select_company(self, company_list):
		self.company_list = company_list
		i=1
		print("Enter a Number to select a Company")
		for comp in self.company_list:
			print(f"[{i}] - {comp.name}")
			i = i+1
		while True:
			try:
				sel = input("Enter Selection: ")
				sel = int(sel)
				selected_company = company_list[int(sel)-1]
				break
			except ValueError: 
				print("Your selection is not a valid integer.  Please try again.")
			except IndexError:
				print("Your selection is not in the correct range.  Please try again.")
		
		return selected_company

