# Deckbot MVP Overview
 
Goal: Given a CompanyID, produce a powerpoint that contains at least 3 slides (cover, company overview, financies [one metric])
	
## MVP Structure
- deckbot.py
  -	Main Application
- models.py
  - Data management, pulling company and financial data from Databook API
- views.py
  - Creates a CLI interface to select a company from a list and sends CompanyID to Presenter
- presenters.py
  - Take CompanyID from view, pull data for that company from Model and generate a powerpoint presentation through python-pptx
- test-presenters.py
  - Unit tests for presenters
- mock-objects.py
  -Mock objects for testing purposes
	
## Dependencies
- requests
- python-pptx
	