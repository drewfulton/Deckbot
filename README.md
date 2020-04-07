# Deckbot MVP Overview
 
Goal: Given a CompanyID, produce a powerpoint that contains at least 3 slides (cover, 
	company overview, financies [one metric])
	
MVP Structure
	deckbot.py
		Main Application
	models.py
		Data management, pulling company and financial data from Databook API
	views.py
		Creates a HTML form with dropdown to select a company and sends CompanyID to Processor
	presenters.py
		Take CompanyID from view, pull data for that company from Model and generate
		a powerpoint presentation through python-pptx
	test_presenters.py
		Unit tests for presenters
	mock_objects.py
		Mock objects for testing purposes
	interactors.py?
		Unknown at the moment
		
		
models.py
	class Company(object)
	class Metric(object)
	
	
	
	GetAllCompanies()
		returns lists of all companies available in Databook
	GetCompany
		returns company object
	GetMetrics(company)
		returns lists of available metrics given a company object
	GetMetric(company, metric)
		returns Metric object given a company and metric name
		
presenters.py
	class DeckbotPresenter(object)
		
	
	
Dependencies
	- requests
	- python-pptx
	