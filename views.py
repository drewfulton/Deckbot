#----------------------------------------------------------------------------
# Name: views.py
# Purpose: 	Provide User Interface
# Author:	Drew Fulton
# Created:	April 4, 2020


class DeckbotCLI(object):
	'''
	This class creates a simple CLI option for creating user input
	'''
	def ___init___(self):
		'''
		Display list of company options and ask for Input from User in CLI
		'''
		self.company_list = company_list
		
	def select_company(self, company_list):
		self.company_list = company_list
		i=1
		print("Enter a Number to select a Company")
		for comp in self.company_list:
			print(f"[{i}] - {comp.name}")
			i = i+1
		sel = input("Enter Selection: ")
		selected_company = company_list[int(sel)-1]
		
# 		print(f"Selected Company: {selected_company.name} - {selected_company.id}")
		return selected_company

